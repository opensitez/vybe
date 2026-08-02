// vybe-test: csharp/common_patterns/fibonacci_iterative
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

int n = 10;
int a = 0, b = 1;
for (int i = 0; i < n; i++) {
    Console.WriteLine(a);
    int tmp = a + b;
    a = b;
    b = tmp;
}
