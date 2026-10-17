use super::*;
use std::{hint::black_box,time::Instant};

fn old_add(a:&[u64],b:&[u64])->Vec<u64> {
    let (long,short)=if a.len()>=b.len(){(a,b)}else{(b,a)};
    let mut out=Vec::with_capacity(long.len()+1);
    let mut carry=0;
    for i in 0..long.len(){
        let (s,c1)=long[i].overflowing_add(*short.get(i).unwrap_or(&0));
        let (s,c2)=s.overflowing_add(carry);
        out.push(s);carry=u64::from(c1)+u64::from(c2);
    }
    if carry>0{out.push(carry);}out
}
fn old_sub(a:&[u64],b:&[u64])->Vec<u64> {
    let mut out=Vec::with_capacity(a.len());let mut borrow=0;
    for i in 0..a.len(){
        let (s,b1)=a[i].overflowing_sub(*b.get(i).unwrap_or(&0));
        let (s,b2)=s.overflowing_sub(borrow);
        out.push(s);borrow=u64::from(b1)+u64::from(b2);
    }
    assert_eq!(borrow,0);mag_norm(out)
}
fn old_or(a:&[u64],b:&[u64])->Vec<u64>{mag_norm((0..a.len().max(b.len())).map(|i|a.get(i).unwrap_or(&0)|b.get(i).unwrap_or(&0)).collect())}
fn old_xor(a:&[u64],b:&[u64])->Vec<u64>{mag_norm((0..a.len().max(b.len())).map(|i|a.get(i).unwrap_or(&0)^b.get(i).unwrap_or(&0)).collect())}
fn old_andnot(a:&[u64],b:&[u64])->Vec<u64>{mag_norm((0..a.len()).map(|i|a[i]&!b.get(i).unwrap_or(&0)).collect())}

type Operation=fn(&[u64],&[u64])->Vec<u64>;
fn methods()->[(&'static str,Operation,Operation);5]{[
    ("add",old_add,mag_add),("sub",old_sub,mag_sub),("or",old_or,mag_or),
    ("xor",old_xor,mag_xor),("andnot",old_andnot,mag_andnot),
]}

#[test]
fn zero_operand_bulk_paths_match_prior_limbs_and_keep_source_independent(){
    for len in 0..129{
        for trailing_zero in [false,true]{
            let mut source:Vec<u64>=(0..len).map(|i|(i as u64).wrapping_mul(0x9e3779b97f4a7c15).wrapping_add(1)).collect();
            if trailing_zero&&len!=0{source[len-1]=0;}
            for (name,old,new) in methods(){
                let mut output=new(&source,&[]);
                assert_eq!(output,old(&source,&[]),"{name} len={len}");
                if let Some(first)=output.first_mut(){*first^=u64::MAX;}
                assert_eq!(new(&source,&[]),old(&source,&[]));
                if matches!(name,"add"|"or"|"xor"){
                    assert_eq!(new(&[],&source),old(&[],&source));
                }
            }
        }
    }
    // Nonempty zero vectors are not canonical zero; old normalization and
    // carry/borrow behavior must still hold without taking the empty shortcut.
    for (_,old,new) in methods(){
        assert_eq!(new(&[u64::MAX,1],&[0,0]),old(&[u64::MAX,1],&[0,0]));
    }
}

#[test]
#[ignore="paired zero-operand typed-buffer bulk-copy benchmark"]
fn zero_operand_bulk_benchmark(){
    for len in [0usize,2,16,1024,64000]{
        let source=vec![u64::MAX;len];
        for zero_rhs in [true,false]{
        if len==0&&!zero_rhs{continue;}
        let right=if zero_rhs{Vec::new()}else{vec![1]};
        for (name,old,new) in methods(){
            let mut samples=[Vec::new(),Vec::new()];
            let rounds=(1048576/len.max(1)).clamp(8,4096);
            for trial in 0..21{
                for order in 0..2{
                    let which=(trial+order)%2;
                    let operation=if which==0{old}else{new};
                    let start=Instant::now();
                    for _ in 0..rounds{black_box(operation(black_box(&source),black_box(&right)));}
                    samples[which].push(start.elapsed().as_nanos() as f64/rounds as f64);
                }
            }
            assert_eq!(old(&source,&right),new(&source,&right));
            for sample in &mut samples{sample.sort_by(f64::total_cmp);}
            println!("zero-operand operation={name} limbs={len} zero_rhs={zero_rhs} old_ns={:.2} bulk_ns={:.2} speedup={:.3}",samples[0][10],samples[1][10],samples[0][10]/samples[1][10]);
        }
        }
    }
}
