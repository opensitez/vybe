// vybe-test: csharp/common_patterns/dictionary_word_count
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

string text = "the cat sat on the mat the cat";
var words = text.Split(' ');
var counts = new Dictionary<string, int>();
foreach (var w in words) {
    if (counts.ContainsKey(w)) counts[w]++;
    else counts[w] = 1;
}
Console.WriteLine("the: " + counts["the"]);
Console.WriteLine("cat: " + counts["cat"]);
Console.WriteLine("sat: " + counts["sat"]);
