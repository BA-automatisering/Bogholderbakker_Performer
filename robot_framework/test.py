class BusinessError(Exception):
    pass

try:
    raise BusinessError("Fejl")
except BusinessError as e:
    print("Fanget:", e)

print("Programmet kører videre")

