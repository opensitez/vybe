// vybe-test: csharp/csharp_random_random/random_next_with_min_max_stays_in_range
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

var rng=new System.Random(1);
for(int i=0;i<100;i++){
    int v=rng.Next(5,10);
    if(v<5||v>=10){Console.WriteLine("fail");return;}
}
Console.WriteLine("pass");
