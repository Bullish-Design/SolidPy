# AUTOGENERATED! DO NOT EDIT! File to edit: ../nbs/02_partdoc.ipynb.

# %% auto 0
__all__ = ['PartDoc']

# %% ../nbs/02_partdoc.ipynb 1
from .modeldoc import ModelDoc
from .interfaces.ipartdoc import IPartDoc


class PartDoc(IPartDoc, ModelDoc):
    def __init__(self, system_object):
        self.system_object = system_object

