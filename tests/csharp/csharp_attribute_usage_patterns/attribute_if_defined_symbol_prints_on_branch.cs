// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_if_defined_symbol_prints_on_branch
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

#define VYBETEST_ON
#if VYBETEST_ON
Console.WriteLine("on");
#else
Console.WriteLine("off");
#endif
