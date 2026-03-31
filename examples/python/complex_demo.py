# Complex Python demo for the minimal compiler
# Computes a 2x2 matrix multiplication (no loops), uses lists, indexing, arithmetic, booleans, and None

# Rows for matrices A and B (2x2)
a0 = [1, 2]
a1 = [3, 4]
b0 = [5, 6]
b1 = [7, 8]

print("debug a0[0]:", a0[0])
print("debug a0[1]:", a0[1])
print("debug b0[0]:", b0[0])
print("debug b1[0]:", b1[0])

# Compute C = A * B (2x2) (flattened)
C00 = a0[0] * b0[0] + a0[1] * b1[0]
C01 = a0[0] * b0[1] + a0[1] * b1[1]
C10 = a1[0] * b0[0] + a1[1] * b1[0]
C11 = a1[0] * b0[1] + a1[1] * b1[1]

print("A:")
print(a0, a1)
print("B:")
print(b0, b1)
print("C (flattened):")
print("C:")
print([C00, C01], [C10, C11])

# Some extra small demos
flag = True
nothing = None
sum0 = C00 + C11
prod0 = C00 * C01
print("flag:")
print(flag)
print("nothing:")
print(nothing)
print("sum0, prod0:")
print(sum0, prod0)
