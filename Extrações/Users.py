from project_requests import InvokeREST

def Usuarios(organization):
    uri = f"https://vssps.dev.azure.com/{organization}/_apis/graph/users?api-version=7.1-preview.1"
    response = InvokeREST(uri)
    return response

def UsuariosPorLicenca(organization):
    uri = f"https://vsaex.dev.azure.com/{organization}/_apis/userentitlements?api-version=7.0"
    response = InvokeREST(uri)
    return response
