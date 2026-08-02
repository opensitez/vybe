// vybe-test: csharp/common_patterns/do_while_loop
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

int x = 1;
do {
    Console.WriteLine(x);
    x *= 3;
} while (x < 100);
