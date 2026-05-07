import os
import requests
from pathlib import Path
from typing import Literal, Union
from io import BytesIO
from datetime import datetime



class SharePointGraph:
    """
    Cliente para integração com SharePoint via Microsoft Graph API.

    Permite realizar operações de gerenciamento de arquivos e pastas em uma
    biblioteca de documentos do SharePoint, utilizando autenticação via
    client credentials (OAuth 2.0).

    Exemplo de uso:
        sp = SharePointGraph(
            tenant_id="seu-tenant-id",
            client_id="seu-client-id",
            client_secret="seu-client-secret",
        )
        sp.download_file(origin_file_path="Pasta/arquivo.pdf", target_path="./downloads")
    """
    @property
    def sharepoint_url(self) -> str:
        """
        Retorna a URL do SharePoint normalizada, sem protocolo e sem barra final.

        Returns:
            str: URL limpa no formato 'dominio.sharepoint.com'.
        """
        # Remove prefixos de protocolo (https:// ou http://)
        sharepoint_url = self.__sharepoint_url.replace("https://", "").replace("http://", "")
        # Remove barra final, se presente
        if sharepoint_url.endswith("/"):
            sharepoint_url = sharepoint_url[:-1]
        return sharepoint_url
    
    def __init__(self, *,
            tenant_id: str,
            client_id: str,
            client_secret: str,
            site_path: str = "rpa",
            sharepoint_url: str = "patrimar.sharepoint.com",
            library_name: str = "Documentos",
        ):
        """
        Inicializa a conexão com o SharePoint via Microsoft Graph API.

        Args:
            tenant_id (str): ID do tenant Azure AD.
            client_id (str): ID do aplicativo registrado no Azure AD.
            client_secret (str): Segredo do aplicativo Azure AD.
            site_path (str): Caminho do site SharePoint (ex: 'rpa'). Padrão: 'rpa'.
            sharepoint_url (str): URL base do SharePoint. Padrão: 'patrimar.sharepoint.com'.
            library_name (str): Nome da biblioteca de documentos. Padrão: 'Documentos'.
        """
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        # URL base da Microsoft Graph API v1.0
        self.base_url = "https://graph.microsoft.com/v1.0"
        # Normaliza o caminho do site removendo prefixos duplicados
        self.site_path = f"sites/{site_path.replace('sites/', '').replace('/', '')}"
        self.__sharepoint_url = sharepoint_url
        self.library_name = library_name

        # Autentica e obtém o token de acesso
        self.token = self._get_token()
        # Cabeçalho de autorização reutilizado em todas as requisições
        self.headers = {"Authorization": f"Bearer {self.token}"}
        # Obtém e armazena o ID do site para uso nas requisições
        self.site_id = self._get_site_id()

    def _get_token(self) -> str:
        """
        Obtém o access token OAuth 2.0 via fluxo client credentials.

        Returns:
            str: Token de acesso Bearer para autenticação nas requisições.

        Raises:
            requests.HTTPError: Se a autenticação falhar.
        """
        # Endpoint de autenticação do Azure AD para o tenant configurado
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        response = requests.post(url, data={
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            # Escopo padrão para acesso à Graph API
            "scope": "https://graph.microsoft.com/.default"
        })
        response.raise_for_status()
        return response.json()["access_token"]

    def _get_site_id(self) -> str:
        """
        Obtém o ID único do site SharePoint na Graph API.

        Returns:
            str: ID do site no formato utilizado pela Graph API.

        Raises:
            requests.HTTPError: Se o site não for encontrado ou o acesso for negado.
        """
        # Monta a URL de consulta combinando a URL do SharePoint com o caminho do site
        url = f"{self.base_url}/sites/{self.sharepoint_url}:/{self.site_path}"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        return response.json()["id"]

    def get_drive_id(self) -> str:
        """
        Obtém o ID do drive correspondente à biblioteca de documentos configurada.

        Returns:
            str: ID do drive da biblioteca.

        Raises:
            ValueError: Se a biblioteca configurada em `library_name` não for encontrada.
            requests.HTTPError: Se a requisição à API falhar.
        """
        url = f"{self.base_url}/sites/{self.site_id}/drives"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        drives = response.json()["value"]

        if self.library_name:
            # Busca a biblioteca pelo nome (comparação case-insensitive)
            for drive in drives:
                if drive["name"].lower() == self.library_name.lower():
                    return drive["id"]
            raise ValueError(f"Biblioteca '{self.library_name}' não encontrada")

        # Se nenhum nome foi configurado, retorna o primeiro drive (geralmente "Documentos")
        return drives[0]["id"]
    
    def list_drivers(self) -> dict[str, str]:
        """
        Lista todas as bibliotecas (drives) disponíveis no site SharePoint.

        Returns:
            dict[str, str]: Dicionário no formato {nome_da_biblioteca: id_do_drive}.
        """
        url = f"{self.base_url}/sites/{self.site_id}/drives"
        response = requests.get(url, headers=self.headers)
        drives = response.json()["value"]

        # Retorna um dicionário mapeando nome -> ID de cada drive
        return {drive['name']: drive['id'] for drive in drives}
    
    def list_root_folders(self) -> list:
        """
        Lista os itens (arquivos e pastas) presentes na raiz da biblioteca.

        Returns:
            list: Lista de dicionários com os metadados de cada item na raiz.
        """
        # Consulta os filhos diretos do nó raiz da biblioteca
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root/children"
        response = requests.get(url, headers=self.headers)
        return response.json()["value"]
    
    def list_folder_items(self, folder_path: str) -> list:
        """
        Lista os itens de uma pasta específica na biblioteca.

        Args:
            folder_path (str): Caminho da pasta relativo à raiz da biblioteca
                               (ex: 'RPA - Dados/Subpasta').
            full_path (bool): Se True, adiciona o campo 'full_path' em cada item
                              com o caminho completo relativo à raiz. Padrão: False.

        Returns:
            list: Lista de dicionários com os metadados de cada item na pasta.
                  Se `full_path=True`, cada item terá o campo extra 'full_path'.
        """
        # Caminho da pasta relativo à raiz da biblioteca
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{folder_path}:/children"
        response = requests.get(url, headers=self.headers)
        items = response.json()["value"]

        # Enriquece cada item com o caminho completo relativo à raiz
        _items = []
        for item in items:
            _items.append(item)
            _items[-1]['full_path'] = f"{folder_path}/{item['name']}"
        items = _items

        return items
    
    def download_file(
            self, 
            origin_file_path: str,
            *,
            save_as: Literal['file', 'binary'],
            target_path: str | Path = "",
            file_name: str = "",
        ) -> BytesIO|str:
        """
        Faz o download de um arquivo do SharePoint.

        Args:
            origin_file_path (str): Caminho do arquivo na biblioteca, relativo à raiz
                                    (ex: 'Pasta/subpasta/arquivo.pdf').
            target_path (str | Path): Diretório local de destino. Se vazio, usa o
                                      diretório de trabalho atual. Ignorado se `save_as='binary'`.
            file_name (str): Nome a ser usado ao salvar o arquivo localmente. Se vazio,
                             mantém o nome original. Ignorado se `save_as='binary'`.
            save_as (Literal['file', 'binary']): Modo de saída.
                - 'file': Salva o arquivo no disco e retorna True.
                - 'binary': Retorna o conteúdo como objeto BytesIO.

        Returns:
            bool: True se o arquivo foi salvo com sucesso (quando `save_as='file'`).
            BytesIO: Conteúdo binário do arquivo (quando `save_as='binary'`).

        Raises:
            NotADirectoryError: Se `target_path` existir mas não for um diretório.
            requests.HTTPError: Se o arquivo não for encontrado ou o acesso for negado.
        """
            # Define o diretório de destino (usa cwd se não informado)
        if not target_path:
            target_path = Path.cwd()
        if isinstance(target_path, str):
            target_path = Path(target_path)
            
        # Requisita o conteúdo binário do arquivo via Graph API
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{origin_file_path}:/content"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
            
        if save_as == 'file':
            # Cria o diretório de destino se não existir
            if not target_path.exists():
                print(f"📁 Criando diretório: {target_path.is_file()=}")
                if target_path.suffix:
                    target_path.parent.mkdir(exist_ok=True)
                else:
                    target_path.mkdir(exist_ok=True)
                    

            # Determina o nome final do arquivo no disco
            if file_name:
                if target_path.suffix:
                    target_path = target_path.parent.joinpath(file_name)
                else:
                    target_path = target_path.joinpath(file_name)
            else:
                if not target_path.suffix:
                    target_path = target_path.joinpath(Path(origin_file_path).name)


        if save_as == 'file':
            # Salva o conteúdo no arquivo local
            with open(str(target_path), "wb") as f:
                f.write(response.content)

            print(f"O Arquivo '{origin_file_path}' foi baixado e salvo em '{target_path.__str__()}'")
            return target_path.__str__()  # Retorna o caminho do arquivo salvo
        elif save_as == 'binary':
            # Retorna o conteúdo em memória como BytesIO
            return BytesIO(response.content)

        raise ValueError("Valor inválido para 'save_as'. Use 'file' ou 'binary'.")
        
    def upload_file(
        self, 
        local_file_path: str | Path,
        *,
        target_path: str | Path = ""
    ) -> bool:
        """
        Faz o upload de um arquivo local para o SharePoint em chunks de 10MB.

        Utiliza o protocolo de upload em sessão da Graph API, adequado para
        arquivos de qualquer tamanho. Arquivos existentes com o mesmo nome
        serão substituídos.

        Args:
            local_file_path (str | Path): Caminho completo do arquivo local a ser enviado.
            target_path (str | Path): Pasta de destino na biblioteca, relativa à raiz
                                      (ex: 'RPA - Dados/Subpasta'). Se vazio, usa a
                                      primeira pasta disponível na raiz.

        Returns:
            bool: True se o upload for concluído com sucesso.

        Raises:
            FileNotFoundError: Se o arquivo local não existir ou não houver pasta de
                               destino disponível na raiz.
            requests.HTTPError: Se alguma requisição à API falhar.
        """
        if isinstance(local_file_path, str):
            local_file_path = Path(local_file_path)
        if not local_file_path.exists():
            raise FileNotFoundError(f"O arquivo '{local_file_path}' não foi encontrado.")
        file_size = os.path.getsize(local_file_path)

        if not target_path:
            # Se nenhum destino for informado, usa a primeira pasta da raiz
            temp = self.list_root_folders()
            if len(temp) > 0:
                target_path = temp[0]['name']
            else:
                raise FileNotFoundError("Não há pastas na raiz da biblioteca para upload. Por favor, crie uma pasta ou especifique um caminho de destino.")

        if isinstance(target_path, str):
            target_path = Path(target_path)
        # Constrói o caminho completo de destino incluindo o nome do arquivo
        if not target_path.suffix:
            target_path = target_path.joinpath(local_file_path.name)
        

        print(target_path)

        # 1. Criar sessão de upload — necessário para envio em múltiplos chunks
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{target_path}:/createUploadSession"
        session = requests.post(url, headers=self.headers, json={
            # Conflito: substitui o arquivo se já existir
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        session.raise_for_status()
        upload_url = session.json()["uploadUrl"]

        # 2. Enviar o arquivo em chunks de 10MB
        chunk_size = 10 * 1024 * 1024  # 10 MB em bytes

        with open(local_file_path, "rb") as f:
            offset = 0
            while offset < file_size:
                chunk = f.read(chunk_size)
                end = offset + len(chunk) - 1

                # Envia o chunk com o cabeçalho Content-Range indicando intervalo de bytes
                response = requests.put(upload_url, headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {offset}-{end}/{file_size}"
                }, data=chunk)
                response.raise_for_status()

                progress = round((end + 1) / file_size * 100, 1)
                print(f"📤 Progresso: {progress}%", end='\r')
                offset += len(chunk)

        print("\n✅ Upload concluído!")
        return True
        
        
    def create_folder(
            self,
            folder_path: str | Path,
            *,
            conflict_behavior: Literal["fail", "rename", "replace"] = "fail"
        ) -> bool:
        """
        Cria uma pasta na biblioteca do SharePoint.

        Args:
            folder_path (str | Path): Caminho da pasta a ser criada, relativo à raiz
                                      (ex: 'RPA - Dados/Nova Pasta').
            conflict_behavior (Literal['fail', 'rename', 'replace']): Comportamento
                em caso de conflito de nome:
                - 'fail': Levanta erro se a pasta já existir (padrão).
                - 'rename': Cria com nome alternativo.
                - 'replace': Substitui a pasta existente.

        Returns:
            bool: True se a pasta for criada com sucesso.

        Raises:
            requests.HTTPError: Se a criação falhar (ex: conflito com 'fail' ou sem permissão).
        """
        if isinstance(folder_path, str):
            folder_path = Path(folder_path)

        # Obtém o caminho do diretório pai para montar a URL correta
        root_path = folder_path.parent.as_posix()

        # Se o pai for a raiz ('.'), usa o endpoint root/children diretamente
        if root_path == ".":
            url = f"{self.base_url}/drives/{self.get_drive_id()}/root/children"
        else:
            url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{root_path}:/children"

        response = requests.post(url, headers={
            **self.headers,
            "Content-Type": "application/json"
        }, json={
            "name": folder_path.name,
            "folder": {},  # Indica que o item a ser criado é uma pasta
            "@microsoft.graph.conflictBehavior": conflict_behavior
        })
        response.raise_for_status()

        print(f"📂 Pasta criada: {folder_path.name}")    
        return True   
    
    def rename_item(self, *, item_path: str | Path, new_name: str) -> bool:
        """
        Renomeia um arquivo ou pasta no SharePoint.

        Args:
            item_path (str | Path): Caminho atual do item relativo à raiz da biblioteca
                                    (ex: 'Pasta/arquivo_antigo.pdf').
            new_name (str): Novo nome para o item (apenas o nome, sem caminho).

        Returns:
            bool: True se o item for renomeado com sucesso.

        Raises:
            requests.HTTPError: Se o item não existir ou o acesso for negado.
        """
        if isinstance(item_path, str):
            item_path = Path(item_path)

        # Primeiro, busca o item pelo caminho para obter seu ID único
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{item_path.as_posix()}"
        item = requests.get(url, headers=self.headers).json()
        item_id = item["id"]

        # Renomeia o item via PATCH usando o ID obtido
        url = f"{self.base_url}/drives/{self.get_drive_id()}/items/{item_id}"
        response = requests.patch(url, headers={
            **self.headers,
            "Content-Type": "application/json"
        }, json={"name": new_name})
        response.raise_for_status()

        print(f"✅ Renomeado para: {response.json()['name']}")
        return True
    
    
    def delete_item(self, *, item_path: str | Path) -> bool:
        """
        Exclui permanentemente um arquivo ou pasta do SharePoint.

        Args:
            item_path (str | Path): Caminho do item a ser excluído, relativo à raiz
                                    da biblioteca (ex: 'Pasta/arquivo.pdf').

        Returns:
            bool: True se o item for excluído com sucesso.

        Raises:
            requests.HTTPError: Se o item não existir ou o acesso for negado.
        """
        if isinstance(item_path, str):
            item_path = Path(item_path)

        # Primeiro, obtém o ID do item pelo caminho
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{item_path.as_posix()}"
        item = requests.get(url, headers=self.headers).json()
        item_id = item["id"]

        # Exclui o item usando o ID obtido
        url = f"{self.base_url}/drives/{self.get_drive_id()}/items/{item_id}"
        response = requests.delete(url, headers=self.headers)
        response.raise_for_status()

        # HTTP 204 (No Content) indica exclusão bem-sucedida
        print(f"🗑️ Arquivo deletado! Status: {response.status_code}")
        return True
    
    def move_item(self, *, item_path: str | Path, target_folder_path: str | Path) -> bool:
        """
        Move um arquivo ou pasta para outro diretório no SharePoint.

        Args:
            item_path (str | Path): Caminho atual do item, relativo à raiz da biblioteca
                                    (ex: 'Pasta/arquivo.pdf').
            target_folder_path (str | Path): Caminho da pasta de destino, relativo à raiz
                                             (ex: 'Outra Pasta/Subpasta').

        Returns:
            bool: True se o item for movido com sucesso.

        Raises:
            requests.HTTPError: Se o item ou a pasta de destino não existirem,
                                ou o acesso for negado.
        """
        if isinstance(item_path, str):
            item_path = Path(item_path)
        if isinstance(target_folder_path, str):
            target_folder_path = Path(target_folder_path)

        # Obtém os IDs do item de origem e da pasta de destino
        item = requests.get(
            f"{self.base_url}/drives/{self.get_drive_id()}/root:/{item_path.as_posix()}", headers=self.headers
        ).json()

        destino_folder = requests.get(
            f"{self.base_url}/drives/{self.get_drive_id()}/root:/{target_folder_path.as_posix()}", headers=self.headers
        ).json()

        # Move o item atualizando sua referência de pai (parentReference)
        url = f"{self.base_url}/drives/{self.get_drive_id()}/items/{item['id']}"
        response = requests.patch(url, headers={
            **self.headers,
            "Content-Type": "application/json"
        }, json={
            "parentReference": {"id": destino_folder["id"]}
        })
        response.raise_for_status()

        print("✅ Arquivo movido!")
        return True
    
    def copy_item(self, *, item_path: str | Path, target_folder_path: str | Path) -> bool:
        """
        Copia um arquivo ou pasta para outro local no SharePoint.

        Se `target_folder_path` tiver uma extensão de arquivo, o nome informado será
        usado como nome da cópia. Caso contrário, o nome original é preservado com
        um sufixo de timestamp no formato `_YYYYMMDDHHmmSS`.

        A operação de cópia é assíncrona na Graph API: a resposta HTTP 202 (Accepted)
        indica que o processo foi iniciado com sucesso em segundo plano.

        Args:
            item_path (str | Path): Caminho do item de origem, relativo à raiz da biblioteca
                                    (ex: 'Pasta/arquivo.pdf').
            target_folder_path (str | Path): Caminho de destino. Se possuir extensão,
                                             trata-se do caminho completo com novo nome
                                             (ex: 'Destino/copia.pdf'). Se for apenas
                                             uma pasta, gera nome com timestamp.

        Returns:
            bool: True se a cópia for iniciada com sucesso (HTTP 202).
                  False se a API retornar outro status.

        Raises:
            requests.HTTPError: Se o item de origem ou a pasta de destino não existirem.
        """
        if isinstance(item_path, str):
            item_path = Path(item_path)
        if isinstance(target_folder_path, str):
            target_folder_path = Path(target_folder_path)

        name = ""
        if target_folder_path.suffix:
            # O destino tem extensão: usa o nome informado como nome da cópia
            name = target_folder_path.name
        else:
            # O destino é uma pasta: gera nome com timestamp para evitar conflito
            sufix = item_path.suffix
            name = item_path.name.replace(sufix, datetime.now().strftime(f"_%Y%m%d%H%M%S{sufix}"))

        # Obtém o ID do item de origem
        item = requests.get(
            f"{self.base_url}/drives/{self.get_drive_id()}/root:/{item_path.as_posix()}", headers=self.headers
        ).json()

        # Obtém o ID da pasta de destino (parent do caminho informado)
        destino_folder = requests.get(
            f"{self.base_url}/drives/{self.get_drive_id()}/root:/{target_folder_path.parent.as_posix()}", headers=self.headers
        ).json()

        # Inicia a cópia assíncrona via Graph API
        url = f"{self.base_url}/drives/{self.get_drive_id()}/items/{item['id']}/copy"
        response = requests.post(url, headers={
            **self.headers,
            "Content-Type": "application/json"
        }, json={
            "parentReference": {"driveId": self.get_drive_id(), "id": destino_folder["id"]},
            "name": name
        })

        if response.status_code == 202:
            # HTTP 202 (Accepted) — cópia iniciada, processamento ocorre em segundo plano
            print(f"✅ Cópia iniciada! Status: {response.status_code}")
            return True

        print(f"❌ Falha ao iniciar cópia. Status: {response.status_code}, Detalhes: {response.text}")
        return False    
    
    

    def item_exists(self, path: str|Path) -> bool:
        """Verifica se um arquivo/pasta existe no SharePoint"""
        
        if isinstance(path, str):
            path = Path(path)
        
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{path.as_posix()}"
        response = requests.get(url, headers=self.headers)
        #response.raise_for_status()
        
        if response.status_code == 200:
            return True
        elif response.status_code == 404:
            return False
        else:
            raise requests.HTTPError(f"Erro ao verificar existência do item. Status: {response.status_code}, Detalhes: {response.text}")
    
    def get_item_metadata(self, path: str|Path):
        """Obtém os metadados de um arquivo/pasta no SharePoint"""
        
        if isinstance(path, str):
            path = Path(path)
        
        url = f"{self.base_url}/drives/{self.get_drive_id()}/root:/{path.as_posix()}"
        response = requests.get(url, headers=self.headers)
        response.raise_for_status()
        
        return response.json()
    
    def is_folder(self, path: str|Path) -> bool:
        """Verifica se o item no caminho especificado é uma pasta"""
        
        metadata = self.get_item_metadata(path)
        return "folder" in metadata
    
    def is_file(self, path: str|Path) -> bool:
        """Verifica se o item no caminho especificado é um arquivo"""
        
        metadata = self.get_item_metadata(path)
        return "file" in metadata
    

if __name__ == "__main__":
    from dotenv import load_dotenv; load_dotenv()
    
    sp = SharePointGraph(
        tenant_id=os.getenv("AZURE_TENANT_ID",""),
        client_id=os.getenv("AZURE_CLIENT_ID",""),
        client_secret=os.getenv("AZURE_CLIENT_SECRET",""),
    )
    
    # buffer = sp.download_file(
    #     origin_file_path=r"RPA - Dados\Relatorio_Imobme_Financeiro\ControleEstoque.json",
    #     target_path=Path.cwd().joinpath("testando", "teste.json"),
    #     save_as='file'
    # )
    # print(buffer)
    
    x = sp.item_exists(r"RPA - Dados/Configs/Vcred/control_files.json")
    
    print(x)
    
