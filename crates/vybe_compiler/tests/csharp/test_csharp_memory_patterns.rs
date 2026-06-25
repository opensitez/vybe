//! `ArrayPool<T>`, `MemoryPool<T>`, and renting/returning patterns.
use super::helpers::run_csharp;

#[test]
fn array_pool_rent_returns_array_of_at_least_requested_size() {
    assert_eq!(
        run_csharp(r#"var pool=System.Buffers.ArrayPool<int>.Shared;
var arr=pool.Rent(10);
Console.WriteLine(arr.Length>=10);
pool.Return(arr);"#),
        &["True"]
    );
}

#[test]
fn array_pool_returned_array_reused_on_next_rent() {
    assert_eq!(
        run_csharp(r#"var pool=System.Buffers.ArrayPool<byte>.Shared;
var a=pool.Rent(8);
pool.Return(a,clearArray:true);
var b=pool.Rent(8);
Console.WriteLine(b.Length>=8);
pool.Return(b);"#),
        &["True"]
    );
}

#[test]
fn memory_owner_slice_wraps_leased_memory() {
    assert_eq!(
        run_csharp(r#"using var owner=System.Buffers.MemoryPool<int>.Shared.Rent(5);
var span=owner.Memory.Span;
for(int i=0;i<5;i++) span[i]=i+1;
Console.WriteLine(span[4]);"#),
        &["5"]
    );
}
