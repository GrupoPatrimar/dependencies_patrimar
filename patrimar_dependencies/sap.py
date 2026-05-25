from .functions import P
import win32com.client
import pywintypes
from functools import wraps
import psutil
import subprocess
from time import sleep
import traceback
import re
import os
import threading
import win32gui
import win32con

# Códigos de erro COM do SAP que indicam desconexão
_SAP_COM_DISCONNECTED_CODES = {
    -2147417848,  # O objeto chamado foi desconectado de seus clientes
    -2147352567,  # Falha de chamada COM genérica
    -2147221020,  # Sintaxe inválida / COM não inicializado
}

def _is_com_disconnect_error(error) -> bool:
    """Verifica se o erro é um erro de desconexão COM do SAP."""
    if isinstance(error, pywintypes.com_error):
        return error.args[0] in _SAP_COM_DISCONNECTED_CODES if error.args else False
    return False

class SAPError(Exception):
    def __init__(self, *args: object) -> None:
        super().__init__(*args)

class FindNewID:
    def __init__(self, connection:win32com.client.CDispatch) -> None:
        """
        Inicializa a classe FindNewID com uma lista de IDs de conexão.

        :param connection: Objeto de conexão do SAP.
        """
        self.__connections:list = []
        for x in range(connection.Children.Count):
            self.__connections.append(connection.Children(x).Id)
            
    def target(self, connection:win32com.client.CDispatch):
        """
        Encontra um novo ID de conexão que não está na lista de conexões existentes.

        :param connection: Objeto de conexão do SAP.
        :return: Índice do novo ID de conexão.
        :raises Exception: Se a sessão não for encontrada.
        """
        for x in range(connection.Children.Count):
            if not connection.Children(x).Id in self.__connections:
                return x
        raise Exception("sessão nao encontrada!")


class SapCrashDialogWatcher:
    """Thread daemon que monitora e fecha dialogs de crash do SAP GUI automaticamente."""

    _SAP_DIALOG_TITLES = (
        "sap logon",
        "sap gui",
        "sap front end",
        "sap gui for windows",
    )

    _CONFIRM_TEXTS = {"sim", "yes", "ok", "ja"}

    def __init__(self, interval: float = 2.0) -> None:
        self._interval = interval
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> None:
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._run, daemon=True, name="SapCrashDialogWatcher")
        self._thread.start()

    def stop(self) -> None:
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=self._interval + 1)

    def __enter__(self) -> "SapCrashDialogWatcher":
        self.start()
        return self

    def __exit__(self, *args) -> None:
        self.stop()

    def _run(self) -> None:
        while not self._stop_event.wait(timeout=self._interval):
            try:
                found_windows: list = []

                def _enum_windows_cb(hwnd, _):
                    try:
                        if not win32gui.IsWindowVisible(hwnd):
                            return True
                        title = win32gui.GetWindowText(hwnd)
                        if not title:
                            return True
                        title_lower = title.lower()
                        if any(t in title_lower for t in SapCrashDialogWatcher._SAP_DIALOG_TITLES):
                            found_windows.append(hwnd)
                    except Exception:
                        pass
                    return True

                win32gui.EnumWindows(_enum_windows_cb, None)

                for hwnd in found_windows:
                    try:
                        confirm_buttons: list = []

                        def _enum_children_cb(child_hwnd, _):
                            try:
                                cls = str(win32gui.GetClassName(child_hwnd)).lower()
                                if "button" in cls:
                                    text = win32gui.GetWindowText(child_hwnd).strip().lower().replace("&", "")
                                    if text in SapCrashDialogWatcher._CONFIRM_TEXTS:
                                        confirm_buttons.append(child_hwnd)
                            except Exception:
                                pass
                            return True

                        win32gui.EnumChildWindows(hwnd, _enum_children_cb, None)

                        if confirm_buttons:
                            target = confirm_buttons[0]
                            try:
                                win32gui.SetForegroundWindow(hwnd)
                            except Exception:
                                pass
                            clicked = False
                            try:
                                win32gui.SendMessage(target, win32con.BM_CLICK, 0, 0)
                                clicked = True
                            except Exception:
                                pass
                            if not clicked:
                                try:
                                    win32gui.PostMessage(target, win32con.BM_CLICK, 0, 0)
                                except Exception:
                                    pass
                            print(P("SapCrashDialogWatcher: dialog SAP detectado e fechado automaticamente.", color='yellow'))
                    except Exception:
                        pass
            except Exception:
                pass


class SAPManipulation():
    @property
    def ambiente(self) -> str|None:
        """
        Retorna o ambiente do SAP.

        :return: Ambiente do SAP.
        """
        return self.__ambiente
    
    @property
    def session(self) -> win32com.client.CDispatch:
        """
        Retorna a sessão atual do SAP.

        :return: Sessão do SAP.
        """
        return self.__session
    
    @session.setter
    def session(self, value):
        self.__session = value
    
    @session.deleter
    def session(self):
        """
        Deleta a sessão atual do SAP.
        """
        try:
            del self.__session
        except:
            pass
        
    @property
    def conn_id(self):
        try:
            if (number:=re.search(r"(?<=con\[)\d+(?=\])", self.connection.Id)):
                return int(number.group())
        except (AttributeError, pywintypes.com_error):
            return None
    
    @property
    def using_active_conection(self) -> bool:
        """
        Retorna se está usando uma conexão ativa.

        :return: True se estiver usando uma conexão ativa, False caso contrário.
        """
        return self.__using_active_conection
    
    @property
    def session_count(self) -> int:
        try:
            return self.connection.Children.Count
        except (AttributeError, pywintypes.com_error):
            return 0
    
    @property
    def list_sessions(self) -> dict:
        """
        Retorna uma lista de sessões ativas do SAP.

        :return: Dicionário com o número de sessões ativas.
        """
        sessions = {}
        try:
            for x in range(self.connection.Children.Count):
                sessions[x] = self.connection.Children(x)
        except (AttributeError, pywintypes.com_error):
            pass
        return sessions
    
    def __del__(self):
        try:
            self._crash_watcher.stop()
        except Exception:
            pass
        try:
            self.fechar_sap(all=True)
        except:
            pass
    
    def __init__(self, *, user:str|None="", password:str|None="", ambiente:str|None="", using_active_conection:bool=False, new_conection=False, connect_in_connId=None) -> None:
        """
        Inicializa a classe SAPManipulation.

        :param user: Usuário do SAP.
        :param password: Senha do SAP.
        :param ambiente: Ambiente do SAP.
        :param using_active_conection: Se está usando uma conexão ativa.
        :param new_conection: Se é uma nova conexão.
        :raises Exception: Se não preencher todos os campos necessários.
        """
        if not using_active_conection and not connect_in_connId:
            if not ((user) and (password) and (ambiente)):
                raise Exception(f"""é necessario preencher todos os campos \n
                                {user=}\n
                                {password=} \n 
                                {ambiente=} \n                            
                                """)
        
        self.__using_active_conection = using_active_conection
        self.__connect_in_connId = connect_in_connId
        self.__user:str|None = user
        self.__password:str|None = password
        self.__ambiente:str|None = ambiente
        self.__new_connection:bool = new_conection
        self._crash_watcher = SapCrashDialogWatcher()
        self._crash_watcher.start()
         
    # Decorador para iniciar o SAP
    @staticmethod
    def start_SAP(f):
        """
        Decorador para iniciar o SAP antes de executar a função decorada.

        :param f: Função a ser decorada.
        :return: Função decorada.
        """
        @wraps(f)
        def wrap(self, *args, **kwargs):
            _self:SAPManipulation = self
            
            # Reconectar quando o COM SAP é desconectado em tempo de execução
            for _reconnect_attempt in range(2):
                try:
                    _self.session
                except AttributeError:
                    _self._SAPManipulation__conectar_sap() #type: ignore
                try:
                    result = f(_self, *args, **kwargs)
                except pywintypes.com_error as com_err:
                    if _is_com_disconnect_error(com_err) and _reconnect_attempt == 0:
                        err_code = com_err.args[0] if com_err.args else None
                        print(P(f"Sessão SAP desconectada (COM error {err_code:#010x}). Reconectando...", color='yellow'))
                        del _self.session
                        _self._SAPManipulation__conectar_sap() #type: ignore
                        continue
                    raise
                else:
                    # Sucesso: delay de estabilização e fechar se solicitado
                    sleep(5)
                    try:
                        if kwargs.get('fechar_sap_no_final'):
                            _self.fechar_sap()
                    except:
                        pass
                    return result
            # Segurança: todas as tentativas falharam sem exceção propagada
            raise SAPError("Não foi possível executar a operação SAP após reconexão")
        return wrap
    
    @staticmethod
    def fechar_sap_no_final(f):
        @wraps(f)
        def wrap(self, *args, **kwargs):
            _self:SAPManipulation = self
            
            result =  f(_self, *args, **kwargs)
            _self.fechar_sap()
            
            return result
        return wrap
    
    # Decorador para verificar conexões
    @staticmethod
    def __verificar_conections(f):
        """
        Decorador para verificar conexões antes de executar a função decorada.

        :param f: Função a ser decorada.
        :return: Função decorada.
        """
        @wraps(f)
        def wrap(self, *args, **kwargs):
            _self:SAPManipulation = self
            
            result = f(_self, *args, **kwargs)
            try:
                if "Continuar com este logon sem encerrar os logons existentes".lower() in (choice:=_self.session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2")).text.lower():
                    choice.select()
                    _self.session.findById("wnd[0]").sendVKey(0)
            except:
                pass
            return result
        return wrap
        
    @__verificar_conections
    def __conectar_sap(self) -> None:
        """
        Conecta ao SAP, abrindo uma nova sessão se necessário.

        :raises Exception: Se não for possível conectar ao SAP.
        :raises ConnectionError: Se ocorrer um erro de conexão.
        """
        for _ in range(2):
            self.__session: win32com.client.CDispatch
            if not self.using_active_conection and not self.__connect_in_connId:
                for tentativa in range(3):
                    try:
                        if not self.__verificar_sap_aberto():
                            #C:\Program Files\SAP\FrontEnd\SAPGUI
                            
                            if os.path.exists(r'C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe'):
                                subprocess.Popen(r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe")
                                
                            elif os.path.exists(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"):
                                subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
                                
                            else:
                                raise Exception("Não foi possível encontrar o executável do SAP Logon. Verifique se o SAP está instalado corretamente.")
                            for _ in range(30):
                                sleep(2)
                                if self.__verificar_sap_aberto():
                                    sleep(5)
                                    break
                        
                        SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
                        self.application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
                        
                        for _ in range(60*60):
                            try:
                                if self.__new_connection:
                                    raise Exception("Erro controlado")
                                
                                conected_info = self.application.Children(0).Children(0).Info
                                if conected_info.SystemName.lower() != self.__ambiente.lower():# type: ignore
                                    raise Exception("Erro controlado")
                                if conected_info.User.lower() != self.__user.lower():# type: ignore
                                    raise Exception("Erro controlado")
                                
                                self.connection = self.application.Children(0) # type: ignore
                            except:
                                self.connection = self.application.OpenConnection(self.__ambiente, True) # type: ignore
                                self.__session = self.connection.Children(0)# type: ignore
                                self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.__user # Usuario
                                self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.__password # Senha
                                self.session.findById("wnd[0]").sendVKey(0)
                                break
                            
                            if _ % 60 == 0 and _ > 0:
                                print(P(f"Aguardando sessão SAP disponível... ({_}s)", color='yellow'))
                                
                            if _ >= ((60*60) - 2):
                                raise Exception("SAP não respondeu após tempo limite. Verifique se há sessões disponíveis.")
                            
                            if self.connection.Children.Count >= 6:
                                sleep(1)
                                continue
                            
                            novo_id = FindNewID(self.connection)
                            session = self.connection.Children(0)# type: ignore
                            
                            session.findById("wnd[0]").sendVKey(74)
                            
                            sleep(1)
                            self.__session = self.connection.Children(novo_id.target(self.connection))# type: ignore
                            break
                                        
                        try:
                            if (sbar:=self.session.findById("wnd[0]/sbar").text):
                                print(P(sbar, color="cyan"))
                        except:
                            pass
                        try:
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press() 
                        except:
                            pass
                        print(P(f"SAP conectado com Sucesso! na tentativa {tentativa + 1}", color='green'))
                        return 

                    except Exception as error:
                        if "sessão nao encontrada!" in str(error):
                            print(P("não foi possivel se conectar a mais uma tela do SAP", color='red'))
                            self.finalizar_programa_sap()
                            continue
                        elif _is_com_disconnect_error(error):
                            if tentativa >= 2:
                                raise Exception("não foi possivel se conectar a mais uma tela do SAP")
                            print(P(f"Erro COM ao conectar ao SAP (tentativa {tentativa + 1}/3). Aguardando...", color='yellow'))
                            sleep(10)
                            continue
                        elif "self.application.OpenConnection" in traceback.format_exc():
                            raise Exception("SAP está fechado!")
                        else:
                            raise ConnectionError(f"não foi possivel se conectar ao SAP motivo: {type(error).__name__} -> {error}")
            else:
                try:
                    if not self.__verificar_sap_aberto():
                        raise Exception("SAP está fechado!")
                    
                    self.SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")
                    self.application: win32com.client.CDispatch = self.SapGuiAuto.GetScriptingEngine

                    if self.__connect_in_connId:
                        try:
                            self.connection: win32com.client.CDispatch = self.application.Children(self._transform_connId_to_connKey(self.__connect_in_connId))
                        except:
                            raise SAPError(f"Não foi possível encontrar a conexão com o id {self.__connect_in_connId}. Verifique se o id está correto e se a conexão está ativa.")
                    else:
                        self.connection: win32com.client.CDispatch = self.application.Children(0)
                    
                    self.__session = self.connection.Children(0)
                    return
                
                except SAPError as error:
                    raise error
                except Exception as error:
                    if "self.connection: win32com.client.CDispatch = self.application.Children(0)" in traceback.format_exc():
                        raise Exception("SAP está fechado!")
                    elif "SAP está fechado!" in traceback.format_exc():
                        raise Exception("SAP está fechado!")
                    else:
                        raise error
        raise Exception("Não foi possível conectar ao SAP após todas as tentativas")
    
    def _transform_connId_to_connKey(self, connId):
        for x in range(self.application.Children.Count):
            if f"con[{connId}]" in self.application.Children(x).Id:
                return x
        raise SAPError("Conexão não encontrada!")
    
    def fechar_sap(self, *, all:bool=False):
        try:
            if not all:
                return self.close_app_sap()
            
            connections_list = [value for value in range(self.connection.Children.Count)]
            connections_list.reverse()
            
            for con in connections_list:
                self.session = self.connection.Children(con)
                self.close_app_sap()
        except (AttributeError, pywintypes.com_error) as error:
            print(P(f"Erro ao fechar SAP: {type(error).__name__} -> {error}", color='red'))
    
    # Método para fechar o SAP
    def close_app_sap(self, *, session_id:int|None=None):
        """
        Fecha a sessão atual do SAP.
        """
        print(P("fechando SAP!", color='red'))
        try:
            session = self.session
        except AttributeError:
            return
        if session_id:
            try:
                session = self.connection.Children(session_id)
            except (AttributeError, pywintypes.com_error):
                return
        try:
            sleep(1)
            session.findById("wnd[0]").close()
            sleep(1)
            try:
                try:
                    session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
                except:
                    session.findById('wnd[2]/usr/btnSPOP-OPTION1').press()
            finally:
                del session
        except Exception as error:
            print(P(f"não foi possivel fechar o SAP {type(error).__name__} | {error}", color='red'))

    # Método para listar elementos
    @start_SAP
    def _listar(self, campo):
        """
        Lista os elementos de um campo específico.

        :param campo: Campo a ser listado.
        """
        cont = 0
        for child_object in self.session.findById(campo).Children:
            print(f"{cont}: ","ID:", child_object.Id, "| Type:", child_object.Type, "| Text:", child_object.Text)
            cont += 1
            
    def create_new_session(self):
        self.session.CreateSession()
        return self.connection.Children.Count
    
    def set_actual_session(self, session_id:int, *, ignore_error:bool=False) -> bool:
        try:
            self.session = self.connection.Children(session_id)
            return True
        except Exception as error:
            if "The enumerator of the collection cannot find an element with the specified index." in traceback.format_exc():
                if ignore_error:
                    print(P(f"Não foi possível encontrar a sessão com o id {session_id}, mas o erro foi ignorado.", color='red'))
                    return False
                raise SAPError(f"Não foi possível encontrar a sessão com o id {session_id}")
            if ignore_error:
                print(P(f"Erro desconhecido ao tentar encontrar a sessão com o id {session_id}, mas o erro foi ignorado. Erro: {type(error).__name__} | {error}", color='red'))
                return False
            raise error
        
    

    def get_actual_session_id(self, session:win32com.client.CDispatch|None=None, *, ignore_error:bool=False) -> int|None:
        if session is None:
            session = self.session
        if isinstance(session, win32com.client.CDispatch):
            if (number:=re.search(r"(?<=ses\[)\d+(?=\])", session.Id)):
                return int(number.group())
        return None

    # Método para verificar se o SAP está aberto
    def __verificar_sap_aberto(self) -> bool:
        """
        Verifica se o SAP está aberto.

        :return: True se o SAP estiver aberto, False caso contrário.
        """
        for process in psutil.process_iter(['name']):
            try:
                if "saplogon" in process.name().lower():
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False    
    
    
    def finalizar_programa_sap(self):
        for proc in psutil.process_iter(['name']):
            try:
                proc_name = proc.info.get('name') or ''
                if "sap" in proc_name.lower():
                    proc.kill()
                    print(P("Processo SAP encerrado.", color='red'))
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue    
    
    # Método de teste         
    @start_SAP
    def _teste(self):
        """
        Método de teste para verificar a conexão com o SAP.
        """
        print("testado")
  
if __name__ == "__main__":
    pass
    #crd = Credential("SAP_QAS").load()
    
    #bot = SAPManipulation(user=crd['user'], password=crd['password'], ambiente="S4Q")
    #bot.conectar_sap()
    #bot._teste()
    
    #import pdb;pdb.set_trace()
    #bot.fechar_sap()
