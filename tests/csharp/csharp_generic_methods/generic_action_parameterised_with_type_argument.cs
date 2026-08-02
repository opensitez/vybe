// vybe-test: csharp/csharp_generic_methods/generic_action_parameterised_with_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

void ForEach<T>(T[] items,System.Action<T> action){
    foreach(var i in items) action(i);
}
ForEach(new[]{1,2,3},n=>Console.WriteLine(n));
