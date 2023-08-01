import settings
from os import makedirs
from sys import argv
from excel import Excel
from TeamProjects import *
from Users import *


def TeamProjects():
    makedirs('TeamProjects', exist_ok=True)
    print('Criando diretório TeamProjects')
    Excel(Projetos(settings.organization))
    Excel(Repositorios(settings.organization))
    Excel(Endpoints(settings.organization))
    Excel(SourceProviders(settings.organization))


def Users():
    makedirs('Users', exist_ok=True)
    print('Criando diretório Users')
    Excel(Usuarios(settings.organization))
    Excel(UsuariosPorLicenca(settings.organization))


def BuildsReleases():
    makedirs('BuildsReleases', exist_ok=True)
    print('Criando diretório BuildsReleases')
    Excel(Builds(settings.organization))
    Excel(Releases(settings.organization))


if __name__ == '__main__':

    parameter = argv[1]
    extracao = eval(parameter)

    extracao()
