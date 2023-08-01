from project_requests import InvokeREST


def Projetos(organization):
    uri = f"https://dev.azure.com/{organization}/_apis/projects?api-version=7.0"
    response = InvokeREST(uri)
    return response

def Repositorios(organization, project = ''):
    uri = f"https://dev.azure.com/{organization}/{project}/_apis/git/repositories?api-version=7.1-preview.1"
    response = InvokeREST(uri)
    return response

def Endpoints(organization, project = ''):
    uri = f"https://dev.azure.com/{organization}/{project}/_apis/serviceendpoint/endpoints?api-version=7.0"
    response = InvokeREST(uri)
    return response

def SourceProviders(organization, project= ''):
    try:
        uri = f"https://dev.azure.com/{organization}/{project}/_apis/sourceproviders?api-version=7.0"
        response = InvokeREST(uri)
        return response
    except:
        None

def Builds(organization, project = ''):
    uri = f"https://dev.azure.com/{organization}/{project}/_apis/build/builds?api-version=7.1-preview.7"
    response = InvokeREST(uri)
    return response

def Releases(organization, project = ''):
    uri = f"https://vsrm.dev.azure.com/{organization}/{project}/_apis/release/releases?api-version=7.1-preview.8"
    response = InvokeREST(uri)
    return response