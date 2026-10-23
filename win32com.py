# Finto modulo win32com per far girare lo script su Linux (GitHub Actions)
class Dispatch:
    def __init__(self, *args, **kwargs): pass
    def __getattr__(self, name): return self
    def __call__(self, *args, **kwargs): return self

class client:
    Dispatch = Dispatch
