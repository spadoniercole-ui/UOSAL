# Finto client per far girare lo script su Linux
class Dispatch:
    def __init__(self, *args, **kwargs): pass
    def __getattr__(self, name):
        def wrapper(*args, **kwargs): return self
        return wrapper
    def __call__(self, *args, **kwargs): return self
