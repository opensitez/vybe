use super::{snapshot_char_code_range, copy_char_code_units};
use crate::Value;
use std::sync::Arc;

#[test]
fn direct_copy_releases_only_overwritten_references() {
    let roots: Vec<Arc<str>> = (0..8).map(|i| Arc::from(format!("root_{i}"))).collect();
    let mut values: Vec<Value> = roots.iter().cloned().map(Value::String).collect();
    assert_eq!(copy_char_code_units("\u{1f980}x", &mut values, 7), None);
    assert!(roots.iter().all(|root| Arc::strong_count(root) == 2));
    for (value, root) in values.iter().zip(&roots) {
        assert!(matches!(value,Value::String(text) if Arc::ptr_eq(text,root)));
    }
    assert_eq!(copy_char_code_units("\u{1f980}x", &mut values, 2), Some(3));
    for (i, root) in roots.iter().enumerate() {
        if (2..5).contains(&i) {
            assert_eq!(Arc::strong_count(root),1);
        } else {
            assert_eq!(Arc::strong_count(root),2);
            assert!(matches!(&values[i],Value::String(text) if Arc::ptr_eq(text,root)));
        }
    }
    drop(values);
    assert!(roots.iter().all(|root| Arc::strong_count(root) == 1));
}

#[test]
fn large_unicode_copies_match_buffered_reference_at_all_offsets() {
    let mut source = String::from("A\0\u{03a9}\u{4e2d}\u{1f980}");
    let mut seed = 0x8a5c_317du32;
    for _ in 0..16384 {
        seed ^= seed << 13;
        seed ^= seed >> 17;
        seed ^= seed << 5;
        if let Some(ch) = char::from_u32(seed % 0x110000) {source.push(ch)}
    }
    let units: Vec<_> = source.encode_utf16().collect();
    for start in 0..16 {
        let mut actual = vec![Value::I32(-1); start+units.len()+16];
        assert_eq!(copy_char_code_units(&source,&mut actual,start),Some(units.len()));
        for (i,value) in actual.iter().enumerate() {
            let expected = if (start..start+units.len()).contains(&i) {units[i-start] as i32} else {-1};
            assert_eq!(value.as_i32(),expected);
        }
        let mut short = vec![Value::I32(-2);start+units.len()-1];
        assert_eq!(copy_char_code_units(&source,&mut short,start),None);
        assert!(short.iter().all(|value|value.as_i32()==-2));
    }
}

#[test]
fn direct_utf16_copy_preserves_units_and_traps_before_mutation() {
    for source in ["", "ASCII\0", "\u{03a9}\u{4e2d}", "\u{1f980}x\u{10400}"] {
        let expected: Vec<u16> = source.encode_utf16().collect();
        for len in 0..=expected.len()+3 {
            for start in 0..=len+1 {
                let mut values = vec![Value::I32(-7);len];
                let result = copy_char_code_units(source,&mut values,start);
                if start <= len && expected.len() <= len-start {
                    assert_eq!(result,Some(expected.len()));
                    for (i,value) in values.iter().enumerate() {
                        let expected = if (start..start+expected.len()).contains(&i) {
                            expected[i-start] as i32
                        } else {-7};
                        assert_eq!(value.as_i32(),expected);
                    }
                } else {
                    assert_eq!(result,None);
                    assert!(values.iter().all(|value| value.as_i32()==-7));
                }
            }
        }
    }
    let mut values = vec![Value::I32(42);4];
    assert_eq!(copy_char_code_units("x",&mut values,usize::MAX),None);
    assert!(values.iter().all(|value| value.as_i32()==42));
}

#[test]
#[ignore = "paired UTF16 write measurement; excludes source/destination allocation"]
fn direct_utf16_copy_benchmark() {
    use std::{hint::black_box,time::Instant};
    for (pattern,repeats) in [("x",1),("abcdefg",1),("\u{1f980}",1),("x",65536),("\u{03a9}",65536),("x\u{1f980}\u{4e2d}",16384)] {
        let source=pattern.repeat(repeats);
        let count=source.encode_utf16().count();
        let mut values=vec![Value::I32(0);count];
        let rounds=if count<16 {4096} else {8};
        let mut samples:[Vec<f64>;2]=std::array::from_fn(|_|Vec::new());
        for trial in 0..21 {
            for new in if trial%2==0 {[false,true]} else {[true,false]} {
                let begin=Instant::now();
                for _ in 0..rounds {
                    if new {
                        assert_eq!(copy_char_code_units(black_box(&source),black_box(&mut values),0),Some(count));
                    } else {
                        let units:Vec<u16>=black_box(&source).encode_utf16().collect();
                        for (i,&unit) in units.iter().enumerate() {
                            black_box(&mut values)[i]=Value::I32(unit as i32);
                        }
                        black_box(units.len());
                    }
                }
                samples[usize::from(new)].push(begin.elapsed().as_secs_f64()/rounds as f64);
                assert!(values.iter().zip(source.encode_utf16()).all(|(v,u)|v.as_i32()==u as i32));
            }
        }
        for sample in &mut samples {sample.sort_by(f64::total_cmp)}
        println!("units={count} bytes={} old={:.3}us direct={:.3}us ratio={:.3}x",source.len(),samples[0][10]*1e6,samples[1][10]*1e6,samples[0][10]/samples[1][10]);
    }
}

#[test]
fn ranges_match_full_snapshot_and_reject_invalid_bounds() {
    let values: Vec<Value> = (0..257).map(|i| Value::I32(i*257-32768)).collect();
    for start in 0..=258 {
        for end in 0..=258 {
            let old = values.clone();
            let expected = old.get(start..end).map(|range|
                range.iter().map(|v| v.as_i32() as u16).collect::<Vec<_>>());
            let actual = snapshot_char_code_range(&values,start,end).map(|range|
                range.iter().map(|v| v.as_i32() as u16).collect::<Vec<_>>());
            assert_eq!(actual,expected);
        }
    }
    assert!(snapshot_char_code_range(&values,usize::MAX,usize::MAX).is_none());
    assert!(snapshot_char_code_range(&values,0,usize::MAX).is_none());
}

#[test]
fn only_selected_references_are_retained_by_snapshot() {
    let first: Arc<str> = Arc::from("first");
    let selected: Arc<str> = Arc::from("selected");
    let last: Arc<str> = Arc::from("last");
    let values = vec![Value::String(first.clone()),Value::String(selected.clone()),Value::String(last.clone())];
    let snapshot = snapshot_char_code_range(&values,1,2).unwrap();
    assert_eq!(Arc::strong_count(&first),2);
    assert_eq!(Arc::strong_count(&selected),3);
    assert_eq!(Arc::strong_count(&last),2);
    assert!(matches!(&snapshot[0],Value::String(value) if Arc::ptr_eq(value,&selected)));
    drop(snapshot);
    assert_eq!(Arc::strong_count(&selected),2);
    assert!(snapshot_char_code_range(&values,3,3).unwrap().is_empty());
}

#[test]
#[ignore = "paired range snapshot measurement; excludes string conversion"]
fn range_snapshot_benchmark() {
    use std::{hint::black_box,time::Instant};
    for (len,width) in [(16usize,1usize),(4096,1),(65536,16),(65536,65536)] {
        let values: Vec<Value> = (0..len).map(|i| Value::I32(i as i32)).collect();
        let start=(len-width)/2;
        let mut samples:[Vec<f64>;2]=std::array::from_fn(|_|Vec::new());
        for trial in 0..21 {
            for new in if trial%2==0 {[false,true]} else {[true,false]} {
                let begin=Instant::now();
                for _ in 0..32 {
                    if new {
                        black_box(snapshot_char_code_range(black_box(&values),start,start+width).unwrap());
                    } else {
                        let copy=black_box(&values).clone();
                        black_box(&copy[start..start+width]);
                    }
                }
                samples[usize::from(new)].push(begin.elapsed().as_secs_f64()/32.);
            }
        }
        for sample in &mut samples {sample.sort_by(f64::total_cmp)}
        println!("array={len} selected={width} full={:.3}us range={:.3}us ratio={:.3}x",samples[0][10]*1e6,samples[1][10]*1e6,samples[0][10]/samples[1][10]);
    }
}
