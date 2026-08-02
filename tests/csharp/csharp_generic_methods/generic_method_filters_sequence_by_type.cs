// vybe-test: csharp/csharp_generic_methods/generic_method_filters_sequence_by_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

System.Collections.Generic.IEnumerable<T> FilterType<T>(object[] items){
    foreach(var i in items) if(i is T t) yield return t;
}
var items=new object[]{1,"a",2,"b",3};
int count=0;
foreach(var s in FilterType<string>(items)) count++;
Console.WriteLine(count);
