// vybe-test: csharp/csharp_ref_out_in/out_parameter_initialised_inside_callee
// origin: languages/csharp/tests/csharp/test_csharp_ref_out_in.rs

void Minmax(int[] a, out int min, out int max){
    min=a[0]; max=a[0];
    foreach(var v in a){if(v<min)min=v; if(v>max)max=v;}
}
Minmax(new[]{3,1,4,1,5,9}, out int lo, out int hi);
Console.WriteLine(lo); Console.WriteLine(hi);
