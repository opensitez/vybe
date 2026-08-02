# vybe-test: python/while_loops/while_reverses_digits_of_number
# origin: languages/python/tests/python/test_while_loops.rs

n = 123
rev = 0
while n:
 rev = rev * 10 + n % 10
 n //= 10
print(rev)
