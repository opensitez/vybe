// vybe-test: csharp/csharp_linq_quantifiers_partition/partition_via_skip_take_page_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

var src=new[]{1,2,3,4,5,6};
int pageSize=2;
int pageCount=0;
for(int i=0;i<src.Length;i+=pageSize) pageCount++;
Console.WriteLine(pageCount);
