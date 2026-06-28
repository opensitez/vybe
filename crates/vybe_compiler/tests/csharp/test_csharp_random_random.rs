//! `System.Random` and `System.Random.Shared` — bounded generation.
use super::helpers::run_csharp;

#[test]
fn random_next_within_exclusive_upper_bound() {
    assert_eq!(
        run_csharp(
            r#"var rng=new System.Random(42);
for(int i=0;i<100;i++){
    int v=rng.Next(10);
    if(v<0||v>=10){Console.WriteLine("fail");return;}
}
Console.WriteLine("pass");"#
        ),
        &["pass"]
    );
}

#[test]
fn random_next_with_min_max_stays_in_range() {
    assert_eq!(
        run_csharp(
            r#"var rng=new System.Random(1);
for(int i=0;i<100;i++){
    int v=rng.Next(5,10);
    if(v<5||v>=10){Console.WriteLine("fail");return;}
}
Console.WriteLine("pass");"#
        ),
        &["pass"]
    );
}

#[test]
fn random_next_double_in_zero_one() {
    assert_eq!(
        run_csharp(
            r#"var rng=new System.Random(7);
for(int i=0;i<100;i++){
    double v=rng.NextDouble();
    if(v<0.0||v>=1.0){Console.WriteLine("fail");return;}
}
Console.WriteLine("pass");"#
        ),
        &["pass"]
    );
}

#[test]
fn seeded_random_produces_deterministic_sequence() {
    assert_eq!(
        run_csharp(
            r#"var r1=new System.Random(99); var r2=new System.Random(99);
Console.WriteLine(r1.Next()==r2.Next());"#
        ),
        &["True"]
    );
}

#[test]
fn random_shuffle_changes_order_of_array() {
    assert_eq!(
        run_csharp(
            r#"int[] arr={1,2,3,4,5,6,7,8,9,10};
var rng=new System.Random(42);
rng.Shuffle(arr);
bool changed=false;
for(int i=0;i<arr.Length;i++) if(arr[i]!=i+1){changed=true;break;}
Console.WriteLine(changed);"#
        ),
        &["True"]
    );
}
