include!("../src/emitter/sdl_runtime.rs");

#[cfg(test)]
mod palette_row_tests {
    use super::*;
    const COLORS: [[u8; 4]; 4] = [[10, 20, 30, 11], [40, 50, 60, 31],
        [70, 80, 90, 51], [100, 110, 120, 71]];

    fn seed(width: usize, rows: usize, pitch: usize, source: usize, first: usize) -> Vec<u8> {
        let mut memory = vec![0xa5; 512];
        for row in 0..rows {
            for x in 0..width { memory[source + row * pitch + x] = (first + (row + x) % 4) as u8; }
        }
        memory
    }

    fn expected(mut memory: Vec<u8>, width: usize, rows: usize, destination: usize, pitch: usize) -> Vec<u8> {
        for row in 0..rows {
            for x in 0..width {
                let target = destination + row * pitch + x * 4;
                memory[target..target + 4].copy_from_slice(&COLORS[(row + x) % 4]);
            }
        }
        memory
    }

    #[test]
    fn complete_palette_covers_every_byte_index_with_exact_alpha_and_padding() {
        let colors: [[u8; 4]; 256] = std::array::from_fn(|i|
            [i as u8, (i * 37) as u8, (i * 13) as u8, (i * 17 + 9) as u8]);
        let mut memory = vec![0xa5; 2700];
        for row in 0..2 {
            for x in 0..256 { memory[8 + row * 260 + x] = ((x + row * 97) % 256) as u8; }
        }
        let mut reference = memory.clone();
        for row in 0..2 {
            for x in 0..256 {
                let index = memory[8 + row * 260 + x] as usize;
                let target = 600 + row * 1040 + x * 4;
                reference[target..target + 4].copy_from_slice(&colors[index]);
            }
        }
        let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
        assert_eq!(expand_linear_palette_rows(&mut ctx, &Value::I32(8), &Value::I32(600),
            0, 0, 256, 2, 260, 1040, 0, &colors), Some(()));
        assert_eq!(memory, reference);
    }

    #[test]
    fn disjoint_expansion_preserves_alpha_padding_offsets_and_both_address_orders() {
        for width in [1usize, 3, 7] {
            for rows in [1usize, 2, 5] {
                for first in [0usize, 3, 252] {
                    for (source, destination) in [(16usize, 200usize), (200, 16)] {
                        let sp = width + 4;
                        let dp = width * 4 + 8;
                        let mut memory = seed(width, rows, sp, source, first);
                        let reference = expected(memory.clone(), width, rows, destination, dp);
                        let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
                        assert_eq!(expand_linear_palette_rows(&mut ctx,
                            &Value::I32((source - 6) as i32), &Value::I32((destination - 6) as i32),
                            6, 6, width, rows, sp, dp, first, &COLORS), Some(()));
                        assert_eq!(memory, reference);
                    }
                }
            }
        }
    }

    #[test]
    fn invalid_final_palette_index_and_overlap_leave_all_memory_unchanged() {
        for invalid in [2u8, 7] {
            let mut memory = seed(7, 5, 11, 16, 3);
            memory[16 + 4 * 11 + 6] = invalid;
            let before = memory.clone();
            let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
            assert_eq!(expand_linear_palette_rows(&mut ctx, &Value::I32(16), &Value::I32(200),
                0, 0, 7, 5, 11, 36, 3, &COLORS), None);
            assert_eq!(memory, before);
        }
        let mut memory = seed(7, 5, 11, 16, 3);
        let before = memory.clone();
        let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
        assert_eq!(expand_linear_palette_rows(&mut ctx, &Value::I32(16), &Value::I32(20),
            0, 0, 7, 5, 11, 36, 3, &COLORS), None);
        assert_eq!(memory, before);
    }

    fn object(fields: &[(&str, Value)]) -> Value {
        let mut value = Object::new();
        value.properties.extend(fields.iter().map(|(name, value)| ((*name).into(), value.clone())));
        Value::Object(heap::alloc(value))
    }

    #[test]
    fn converter_disjoint_fast_path_and_overlapping_snapshot_fallback_agree() {
        let mut colors = Object::new();
        colors.kind = ObjectKind::Array(COLORS.iter().map(|color| object(&[
            ("r", Value::I32(color[0] as i32)), ("g", Value::I32(color[1] as i32)),
            ("b", Value::I32(color[2] as i32)), ("a", Value::I32(color[3] as i32)),
        ])).collect());
        let palette = object(&[("colors", Value::Object(heap::alloc(colors))),
            ("ncolors", Value::I32(4)), ("__sdl_first", Value::I32(3))]);
        let indexed = object(&[("BitsPerPixel", Value::I32(8)), ("palette", palette),
            ("Rmask", Value::I32(0)), ("Gmask", Value::I32(0)),
            ("Bmask", Value::I32(0)), ("Amask", Value::I32(0))]);
        let rgba = object(&[("BitsPerPixel", Value::I32(32)),
            ("Rmask", Value::I64(255)), ("Gmask", Value::I64(65280)),
            ("Bmask", Value::I64(16711680)), ("Amask", Value::I64(4278190080))]);
        for destination in [20usize, 200] {
            let source = object(&[("format", indexed.clone()), ("w", Value::I32(7)),
                ("h", Value::I32(5)), ("pitch", Value::I32(11)), ("pixels", Value::I32(16))]);
            let target = object(&[("format", rgba.clone()), ("w", Value::I32(7)),
                ("h", Value::I32(5)), ("pitch", Value::I32(36)),
                ("pixels", Value::I32(destination as i32))]);
            let mut memory = seed(7, 5, 11, 16, 3);
            let reference = expected(memory.clone(), 7, 5, destination, 36);
            let mut ctx = HostContext::empty(); ctx.memory = Some(&mut memory);
            assert_eq!(convert_blit(&mut ctx, &[source, Value::Null, target, Value::Null, Value::I32(0)]), Some(()));
            assert_eq!(memory, reference);
        }
    }
}
