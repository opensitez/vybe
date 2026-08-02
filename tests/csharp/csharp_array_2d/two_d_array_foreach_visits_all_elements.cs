// vybe-test: csharp/csharp_array_2d/two_d_array_foreach_visits_all_elements
// origin: languages/csharp/tests/csharp/test_csharp_array_2d.rs

int[,] m={{1,2},{3,4}};
int sum=0; foreach(int n in m) sum+=n;
Console.WriteLine(sum);
