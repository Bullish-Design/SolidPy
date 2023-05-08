# AUTOGENERATED! DO NOT EDIT! File to edit: ../nbs/10_component.ipynb.

# %% auto 0
__all__ = ['Component']

# %% ../nbs/10_component.ipynb 1
from .interfaces.icomponent import IComponent


class Component(IComponent):
    def __init__(self, system_object):
        super().__init__(system_object)

    def __repr__(self):
        return f"{self.__class__.__name__} <{self.name}>"

    def __str__(self):
        return self.name

    def get_children(self):
        return [
            Component(system_object) for system_object in self._get_children()
        ]

    def get_parent(self):
        return Component(self._get_parent())

