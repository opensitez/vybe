// vybe-test: csharp/common_patterns/for_loop_with_break_continue
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

for (int i = 0; i < 10; i++) {
    if (i % 2 == 0) continue;
    if (i > 7) break;
    Console.WriteLine(i);
}
