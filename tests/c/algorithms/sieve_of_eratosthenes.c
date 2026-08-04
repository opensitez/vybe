// vybe-test: c/algorithms/sieve_of_eratosthenes
// origin: languages/c/tests/c/test_algorithms.rs
#include <stdio.h>
#include <string.h>
#include <assert.h>

int is_prime[20];
void sieve(int n) {
    for (int i = 0; i < n; i++) is_prime[i] = 1;
    is_prime[0] = is_prime[1] = 0;
    for (int i = 2; i*i < n; i++)
        if (is_prime[i])
            for (int j = i*i; j < n; j += i) is_prime[j] = 0;
}
int main() {
const char *__w[] = {"8\n"};
int __n = 1, __i = 0;
sieve(20);
int count = 0;
for (int i = 0; i < 20; i++) if (is_prime[i]) count++;
{ char __t[512]; snprintf(__t, sizeof(__t), "%d\n", count);
  if (__i >= __n || strcmp(__t, __w[__i]) != 0) { printf("FAIL at line %d: got [%s]\n", __i, __t); assert(0); } __i++; }
if (__i != __n) { printf("FAIL: %d line(s), wanted %d\n", __i, __n); assert(0); }
return 0;
}

