use vybe_ast::{Span, line_index::{self, LineIndex}};

fn reference(source: &str, offset: usize) -> (u32, u32) {
    let mut line = 1u32;
    let mut start = 0;
    for (i, &byte) in source.as_bytes().iter().enumerate() {
        if byte == b'\n' && i < offset {
            line += 1;
            start = i + 1;
        }
    }
    (line, (offset - start) as u32 + 1)
}
fn fields(span: Option<Span>) -> Option<(u32, u32, u32, u32)> {
    span.map(|s| (s.start_line, s.start_col, s.end_line, s.end_col))
}
fn expected(source: &str, start: usize, end: usize, base: u32) -> Option<(u32, u32, u32, u32)> {
    let (sl, sc) = reference(source, start);
    let (el, ec) = reference(source, end);
    Some((sl - base, sc - base, el - base, ec - base))
}
fn check(source: &str, start: usize, end: usize) {
    assert_eq!(line_index::line_col(start), Some(reference(source, start)));
    assert_eq!(fields(line_index::span_0based(start, end)), expected(source, start, end, 1));
    assert_eq!(fields(line_index::span_1based(start, end)), expected(source, start, end, 0));
}

#[test]
fn guard_keeps_original_storage_traits_and_absent_index_behavior() {
    fn send_sync<T: Send + Sync>() {}
    send_sync::<LineIndex>();
    assert_eq!(std::mem::size_of::<LineIndex>(), std::mem::size_of::<Vec<usize>>());
    assert!(line_index::line_col(0).is_none());
    assert!(line_index::span_0based(2, 0).is_none());
    assert!(line_index::span_1based(0, 2).is_none());
    {
        let _outer = LineIndex::install("a\nb");
        check("a\nb", 2, 0);
        {
            let _inner = LineIndex::install("");
            check("", 4, 0);
        }
        check("a\nb", 2, 0);
    }
    assert!(line_index::line_col(0).is_none());
}

#[test]
fn all_byte_endpoints_match_linear_reference() {
    for source in ["", "x", "\n", "\n\n", "a\r\nb\rc\nd", "雪\n🐍é\r\n", "one\ntwo\nthree\n"] {
        let _index = LineIndex::install(source);
        // Byte offsets inside a multibyte scalar and after EOF are intentional:
        // this API numbers bytes and does not validate source boundaries.
        for start in 0..source.len() + 4 {
            for end in 0..source.len() + 4 {
                check(source, start, end);
            }
        }
    }
}

#[test]
fn randomized_nested_sources_restore_every_outer_index() {
    fn next(seed: &mut u64) -> usize {
        *seed = seed.wrapping_mul(6364136223846793005).wrapping_add(1);
        (*seed >> 32) as usize
    }
    let mut seed = 0x51f36a42u64;
    for _ in 0..64 {
        let mut sources = Vec::new();
        for _ in 0..4 {
            let mut source = String::new();
            for _ in 0..next(&mut seed) % 64 {
                source.push_str(["a", "雪", "🐍", "\n", "\r", "\r\n", "\0"][next(&mut seed) % 7]);
            }
            sources.push(source);
        }
        let mut guards = Vec::new();
        for source in &sources {
            guards.push(LineIndex::install(source));
            for _ in 0..32 {
                check(source, next(&mut seed) % (source.len() + 8), next(&mut seed) % (source.len() + 8));
            }
        }
        while !guards.is_empty() {
            guards.pop();
            if let Some(source) = guards.len().checked_sub(1).map(|i| &sources[i]) {
                for _ in 0..32 {
                    check(source, next(&mut seed) % 512, next(&mut seed) % 512);
                }
            } else {
                assert!(line_index::line_col(0).is_none());
            }
        }
    }
}

#[test]
fn far_outside_offsets_preserve_column_cast_and_panic_contract() {
    let mut offsets = vec![0, 1, 10, u32::MAX as usize - 1, u32::MAX as usize, usize::MAX];
    if let Some(after_u32) = (u32::MAX as usize).checked_add(1) {
        offsets.push(after_u32);
    }
    for source in ["", "a\nb\n", "雪\n"] {
        let _index = LineIndex::install(source);
        for &start in &offsets {
            assert_eq!(std::panic::catch_unwind(|| line_index::line_col(start)).ok(),
                       std::panic::catch_unwind(|| Some(reference(source, start))).ok());
            for &end in &[0, start, usize::MAX] {
                assert_eq!(std::panic::catch_unwind(|| fields(line_index::span_0based(start, end))).ok(),
                           std::panic::catch_unwind(|| expected(source, start, end, 1)).ok());
                assert_eq!(std::panic::catch_unwind(|| fields(line_index::span_1based(start, end))).ok(),
                           std::panic::catch_unwind(|| expected(source, start, end, 0)).ok());
            }
        }
    }
}
