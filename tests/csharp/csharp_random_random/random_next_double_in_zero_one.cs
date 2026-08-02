// vybe-test: csharp/csharp_random_random/random_next_double_in_zero_one
// origin: languages/csharp/tests/csharp/test_csharp_random_random.rs

var rng=new System.Random(7);
for(int i=0;i<100;i++){
    double v=rng.NextDouble();
    if(v<0.0||v>=1.0){Console.WriteLine("fail");return;}
}
Console.WriteLine("pass");
