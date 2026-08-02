// vybe-test: csharp/csharp_nullable_ref_types/nullable_reference_type_array_element_can_be_null
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

string?[] arr=new string?[3];
arr[0]="a"; arr[1]=null; arr[2]="c";
int nonNull=0;
foreach(var s in arr) if(s!=null) nonNull++;
Console.WriteLine(nonNull);
