// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_if_else_nested_structural
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

#define VYBETEST_A
#if VYBETEST_A
Console.WriteLine("a");
#else
Console.WriteLine("b");
#endif
Console.WriteLine("c");
