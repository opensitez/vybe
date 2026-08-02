// vybe-test: csharp/csharp_linq_quantifiers_partition/partition_manual_window_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

var src=new[]{10,20,30,40,50};
int size=2;
int windows=0;
for(int i=0;i+size<=src.Length;i+=size) windows++;
Console.WriteLine(windows);
