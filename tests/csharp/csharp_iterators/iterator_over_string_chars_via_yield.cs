// vybe-test: csharp/csharp_iterators/iterator_over_string_chars_via_yield
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

System.Collections.Generic.IEnumerable<char> Vowels(string s) {
    foreach(char c in s) if("aeiou".Contains(c)) yield return c;
}
int count=0;
foreach(var _ in Vowels("hello world")) count++;
Console.WriteLine(count);
