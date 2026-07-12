//! `System.Threading.Tasks.Parallel`: For, ForEach, Invoke.
use super::helpers::run_csharp;

#[test]
fn parallel_for_accumulates_all_iterations() {
    assert_eq!(
        run_csharp(
            r#"int sum=0;
System.Threading.Tasks.Parallel.For(0,100,i=>{
    System.Threading.Interlocked.Add(ref sum,i);
});
Console.WriteLine(sum);"#
        ),
        &["4950"]
    );
}

#[test]
fn parallel_for_each_processes_all_elements() {
    assert_eq!(
        run_csharp(
            r#"var items=new[]{1,2,3,4,5};
int sum=0;
System.Threading.Tasks.Parallel.ForEach(items,n=>{
    System.Threading.Interlocked.Add(ref sum,n);
});
Console.WriteLine(sum);"#
        ),
        &["15"]
    );
}

#[test]
fn parallel_invoke_runs_all_actions() {
    assert_eq!(
        run_csharp(
            r#"int a=0,b=0;
System.Threading.Tasks.Parallel.Invoke(
    ()=>System.Threading.Interlocked.Exchange(ref a,1),
    ()=>System.Threading.Interlocked.Exchange(ref b,2)
);
Console.WriteLine(a+b);"#
        ),
        &["3"]
    );
}
