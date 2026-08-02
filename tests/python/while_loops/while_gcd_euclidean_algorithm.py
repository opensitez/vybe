# vybe-test: python/while_loops/while_gcd_euclidean_algorithm
# origin: languages/python/tests/python/test_while_loops.rs

a = 48
b = 18
while b:
 a, b = b, a % b
print(a)
