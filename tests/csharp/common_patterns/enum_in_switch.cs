// vybe-test: csharp/common_patterns/enum_in_switch
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

enum Season { Spring, Summer, Autumn, Winter }
Season s = Season.Summer;
switch (s) {
    case Season.Spring: Console.WriteLine("spring"); break;
    case Season.Summer: Console.WriteLine("summer"); break;
    case Season.Autumn: Console.WriteLine("autumn"); break;
    case Season.Winter: Console.WriteLine("winter"); break;
}
