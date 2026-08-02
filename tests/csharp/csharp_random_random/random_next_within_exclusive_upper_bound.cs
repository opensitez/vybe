// vybe-test: csharp/csharp_random_random/random_next_within_exclusive_upper_bound
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

var rng=new System.Random(42);
for(int i=0;i<100;i++){
    int v=rng.Next(10);
    if(v<0||v>=10){Console.WriteLine("fail");return;}
}
Console.WriteLine("pass");
