# Test arithmetic, indexing and prints (matrix 2x2 example assertions)
a0 = [1, 2]
a1 = [3, 4]
b0 = [5, 6]
b1 = [7, 8]

C00 = a0[0] * b0[0] + a0[1] * b1[0]
C01 = a0[0] * b0[1] + a0[1] * b1[1]
C10 = a1[0] * b0[0] + a1[1] * b1[0]
C11 = a1[0] * b0[1] + a1[1] * b1[1]

ok = True
if not (C00 == 19 and C01 == 22 and C10 == 43 and C11 == 50):
  ok = False

if ok:
  print("PASS test_arithmetic")
else:
  print("FAIL test_arithmetic")
