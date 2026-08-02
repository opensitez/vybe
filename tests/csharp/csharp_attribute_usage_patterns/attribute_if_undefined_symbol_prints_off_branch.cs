// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_if_undefined_symbol_prints_off_branch
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

#if VYBETEST_OFF
Console.WriteLine("on");
#else
Console.WriteLine("off");
#endif
