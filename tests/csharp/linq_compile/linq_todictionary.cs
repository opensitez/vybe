// vybe-test: csharp/linq_compile/linq_todictionary
// origin: languages/csharp/tests/csharp/test_linq_compile.rs
// vybe-test-mode: compile

var words = new List<string>();
words.Add("a");
words.Add("bb");
var dict = words.ToDictionary(w => w);
