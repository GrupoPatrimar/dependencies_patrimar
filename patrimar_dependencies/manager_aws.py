import boto3

from typing import Literal, List
from pathlib import Path


class ManagerAWS:
    @property
    def s3_client(self):
        return self.__s3_client
    
    @staticmethod
    def fix_prefixo(func):
        def wrapper(self, *args, **kwargs):
            prefixo = kwargs.get('prefixo', '')
            if prefixo and not prefixo.endswith('/'):
                kwargs['prefixo'] = prefixo + '/'
            return func(self, *args, **kwargs)
        return wrapper
    
    def __init__(
        self, *,
        aws_access_key_id:str,
        aws_secret_access_key:str,
        region_name:Literal['us-east-1']|str = 'us-east-1',
        service_name:str = 's3'
    ) -> None:
        

        
        self.__s3_client = boto3.client(
            service_name,
            aws_access_key_id=aws_access_key_id,
            aws_secret_access_key=aws_secret_access_key,
            region_name=region_name
        )
        
    @fix_prefixo
    def list_files(self, bucket_name:str, *, prefixo='') -> list:
        """Lista todos os objetos em um bucket (ou 'pasta')."""   
        
        paginator = self.s3_client.get_paginator('list_objects_v2')
        pages = paginator.paginate(Bucket=bucket_name, Prefix=prefixo, Delimiter='/')
        
        #import pdb; pdb.set_trace(header="Teste de breakpoint")
        
        arquivos = []
        for page in pages:
            for obj in page.get('Contents', []):
                print(obj['Key'])
                name_temp = Path(obj['Key'])
                
                if len(name_temp.parents) > 2:
                    continue
                else:
                    obj['item'] = name_temp.name
                    arquivos.append(obj)
        
        return arquivos
    
    @fix_prefixo
    def list_folders(self, bucket_name:str, *, prefixo='') -> list:
        """Lista todas as 'pastas' em um bucket."""
        
        response = self.s3_client.list_objects_v2(Bucket=bucket_name, Prefix=prefixo, Delimiter='/')
        
        pastas = []
        
        for prefix in response.get('CommonPrefixes', []):
            prefix['item'] = prefix['Prefix']
            pastas.append(prefix)

        return pastas
    
    @fix_prefixo
    def walk(self, bucket_name:str, *, prefixo='', nivel:int=0) -> list:
        """Percorre recursivamente um bucket, listando arquivos e 'pastas'."""
                
        response = self.s3_client.list_objects_v2(Bucket=bucket_name, Prefix=prefixo, Delimiter='/')

        indent = "  " * nivel
    
        
        itens:List[str] = []
        
        # Mostra arquivos do nível atual
        for obj in response.get('Contents', []):
            if obj['Key'] != prefixo:
                #nome = obj['Key'].split('/')[-1]
                obj['item'] = obj['Key']
                itens.append(obj)
                
        
        # Navega nas subpastas
        for prefix in response.get('CommonPrefixes', []):
            nome_pasta = prefix['Prefix'].rstrip('/').split('/')[-1]
            #itens.append(f"{indent}{nome_pasta}/")
            prefix['item'] = prefix['Prefix']
            itens.append(prefix)
            itens.extend(self.walk(bucket_name, prefixo=prefix['Prefix'], nivel=nivel+1))
            
        #return [item.strip() for item in itens]
        return itens
    
    def get_item(self, bucket_name:str, *, key:str):
        """Obtém um objeto específico do bucket."""
        try:
            response = self.s3_client.get_object(Bucket=bucket_name, Key=key)
            return response
        except self.s3_client.exceptions.NoSuchKey:
            print(f"Objeto '{key}' não encontrado no bucket '{bucket_name}'.")
            return None
        
    def item_exists(self, bucket_name:str, *, key:str) -> bool:
        """Verifica se um objeto existe no bucket."""
        try:
            self.s3_client.head_object(Bucket=bucket_name, Key=key)
            return True
        except self.s3_client.exceptions.ClientError as e:
            if e.response['Error']['Code'] == '404':
                return False
            else:
                raise
        
    
if __name__ == "__main__":
    import os
    from dotenv import load_dotenv; load_dotenv()
    
    manager_aws = ManagerAWS(
        aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID', ""),
        aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY', "")
    )
    

    x =manager_aws.item_exists('patrimar', key="BLIP/29-12-2025/15-14-38-36b8f95d-e7ab-4ed1-a218-019b2826bf46.pdf")#, full_path=True)
    
    print(x)

