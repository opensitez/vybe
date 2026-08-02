// vybe-test: csharp/csharp_covariance_contravariance/ienumerable_covariance_allows_derived_sequence_in_base_reference
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

System.Collections.Generic.IEnumerable<string> strings =
    new System.Collections.Generic.List<string> { "x" };
System.Collections.Generic.IEnumerable<object> objects = strings;
foreach (var o in objects) Console.WriteLine(o);
