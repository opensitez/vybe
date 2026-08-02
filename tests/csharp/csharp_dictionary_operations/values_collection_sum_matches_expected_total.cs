// vybe-test: csharp/csharp_dictionary_operations/values_collection_sum_matches_expected_total
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",3},{"b",7}};
int sum=0; foreach(var v in d.Values) sum+=v;
Console.WriteLine(sum);
