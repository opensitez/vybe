//! Operand-size equivalence and performance gates, not guest execution timing.
use vybe_runtime::opcode::{read_leb_u32, read_leb_u64, simd_memarg_size};

fn previous_size(code: &[u8], operand_start: usize) -> usize {
    let mut ip = operand_start;
    let align = read_leb_u32(code, &mut ip);
    if align & 0x80 == 0 {
        return 0;
    }
    let _offset = read_leb_u64(code, &mut ip);
    if align & 0x40 != 0 {
        let _memidx = read_leb_u32(code, &mut ip);
    }
    ip.saturating_sub(operand_start)
}

fn leb(mut value: u64, out: &mut Vec<u8>) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 { byte |= 0x80; }
        out.push(byte);
        if value == 0 { break; }
    }
}

fn marked(align: u32, offset: u64, memidx: Option<u32>) -> Vec<u8> {
    let mut out = Vec::new();
    leb(align as u64, &mut out);
    leb(offset, &mut out);
    if let Some(index) = memidx { leb(index as u64, &mut out); }
    out
}

// Also callable by a tiny standalone harness using the actual opcode source,
// avoiding a second Cargo target directory during other agents' builds.
pub(crate) fn check_operand_size_equivalence() {
    let mut corpus = vec![vec![], vec![0], vec![0, 0xfd, 0, 0x0c], vec![0x80, 0]];
    for lane in 0..16 { corpus.push(vec![lane, 0, 0xfd]); }
    for align in [0x80u32, 0x84, 0xc4, 0x184, 0x1c4] {
        for offset in [0, 1, 127, 128, u32::MAX as u64, u64::MAX] {
            let indices: &[u32] = if align & 0x40 != 0 { &[0, 127, 128, u32::MAX] } else { &[0] };
            for &index in indices {
                let bytes = marked(align, offset, (align & 0x40 != 0).then_some(index));
                assert_eq!(simd_memarg_size(&bytes, 0), bytes.len());
                for length in 0..=bytes.len() { corpus.push(bytes[..length].to_vec()); }
            }
        }
    }
    for code in corpus {
        for start in 0..=code.len() + 2 {
            assert_eq!(simd_memarg_size(&code, start), previous_size(&code, start),
                "code={code:?}, start={start}");
        }
        // An out-of-range host cursor is not a valid instruction position.
        // The existing LEB reader panics on cursor overflow in debug builds
        // and wraps in optimized builds; compare behavior without treating
        // either mode's legacy behavior as a WASM guest guarantee.
        let actual = std::panic::catch_unwind(|| simd_memarg_size(&code, usize::MAX));
        let previous = std::panic::catch_unwind(|| previous_size(&code, usize::MAX));
        match (actual, previous) {
            (Ok(actual), Ok(previous)) => assert_eq!(actual, previous),
            (Err(_), Err(_)) => {},
            _ => panic!("out-of-range cursor behavior differs"),
        }
    }
    // Any first byte without continuation must decode below the marker bit.
    for first in 0..128u8 {
        let code = [first, 0xff, 0xff, 0xff, 0xff, 0xff];
        assert_eq!(simd_memarg_size(&code, 0), previous_size(&code, 0));
    }
}

#[test]
fn optional_memory_operand_sizes_match_previous_decoder() {
    check_operand_size_equivalence();
}

pub(crate) fn measure_operand_sizes() {
    use std::hint::black_box;
    use std::time::Instant;
    type Sizer = fn(&[u8], usize) -> usize;
    fn time(sizer: Sizer, code: &[u8]) -> f64 {
        let start = Instant::now();
        for _ in 0..10_000_000 {
            black_box(black_box(sizer)(black_box(code), black_box(0)));
        }
        start.elapsed().as_secs_f64()
    }
    let cases = [
        ("omitted", vec![0, 0xfd, 0, 0x0c]),
        ("lane", vec![15, 0, 0xfd]),
        ("marked32", marked(0x84, 128, None)),
        ("marked64-multi", marked(0x1c4, u64::MAX, Some(257))),
    ];
    for (name, code) in cases {
        assert_eq!(simd_memarg_size(&code, 0), previous_size(&code, 0));
        let mut ratios = [0.0f64; 5];
        for (pair, ratio) in ratios.iter_mut().enumerate() {
            let (old, new) = if pair % 2 == 0 {
                (time(previous_size, &code), time(simd_memarg_size, &code))
            } else {
                let new = time(simd_memarg_size, &code);
                (time(previous_size, &code), new)
            };
            *ratio = old / new;
        }
        ratios.sort_by(f64::total_cmp);
        eprintln!("memarg size {name}: previous/current median {:.3}x; pairs {ratios:?}", ratios[2]);
    }
}

#[test]
#[ignore = "paired release operand-size helper measurement"]
fn operand_size_benchmark() {
    measure_operand_sizes();
}
