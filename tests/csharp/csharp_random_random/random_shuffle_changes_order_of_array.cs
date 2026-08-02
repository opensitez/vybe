// vybe-test: csharp/csharp_random_random/random_shuffle_changes_order_of_array
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

int[] arr={1,2,3,4,5,6,7,8,9,10};
var rng=new System.Random(42);
rng.Shuffle(arr);
bool changed=false;
for(int i=0;i<arr.Length;i++) if(arr[i]!=i+1){changed=true;break;}
Console.WriteLine(changed);
