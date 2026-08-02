// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_nested_string_arrays_joined_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{new[]{"a","b"},new[]{"c"}}.SelectMany(x=>x);
__Check((flat.Count()).ToString(), "3");
