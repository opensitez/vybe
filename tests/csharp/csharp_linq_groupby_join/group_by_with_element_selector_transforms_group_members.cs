// vybe-test: csharp/csharp_linq_groupby_join/group_by_with_element_selector_transforms_group_members
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

var nums = new[] { 1, 2, 3, 4 };
var groups = nums.GroupBy(n => n % 2 == 0 ? "even" : "odd",
                          n => n * 10);
int evenSum = 0;
foreach (var g in groups)
    if (g.Key == "even") foreach (var v in g) evenSum += v;
Console.WriteLine(evenSum);
