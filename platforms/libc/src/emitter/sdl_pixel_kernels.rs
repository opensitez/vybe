//! Portable pixel kernels over independently owned framebuffer bytes.

/// Map whole-byte color channels after the caller validates the pixel formats.
/// Missing RGB and unused output bytes are zero; missing stored alpha is 255.
pub(crate) fn byte_channel_mapping(
    source: [u32; 4], destination: [u32; 4],
) -> ([Option<usize>; 4], [u8; 4]) {
    let mut mapping = [None; 4];
    let mut defaults = [0; 4];
    for component in 0..4 {
        if destination[component] != 0 {
            let byte = (destination[component].trailing_zeros() / 8) as usize;
            if source[component] != 0 {
                mapping[byte] = Some((source[component].trailing_zeros() / 8) as usize);
            } else if component == 3 {
                defaults[byte] = 255;
            }
        }
    }
    (mapping, defaults)
}

/// Swap red/blue in complete four-byte pixels, preserving green and alpha.
/// Decode words explicitly as little-endian, independently of host byte order.
/// Partial trailing pixels are left untouched.
pub(crate) fn swap_bgra_bytes(bytes: &mut [u8]) {
    let mut words = bytes.chunks_exact_mut(8);
    for word in &mut words {
        let bits = u64::from_le_bytes((&*word).try_into().unwrap());
        let swapped = (bits & 0xff00_ff00_ff00_ff00)
            | ((bits & 0x0000_00ff_0000_00ff) << 16)
            | ((bits & 0x00ff_0000_00ff_0000) >> 16);
        word.copy_from_slice(&swapped.to_le_bytes());
    }
    for pixel in words.into_remainder().chunks_exact_mut(4) {
        pixel.swap(0, 2);
    }
}

/// Expand complete RGB/BGR triples to owned RGBA after validating all rows.
pub(crate) fn expand_rgb24_rows(
    input: &[u8], width: usize, rows: usize, pitch: usize, bgr: bool,
) -> Option<Vec<u8>> {
    let row_bytes = width.checked_mul(3)?;
    let length = width.checked_mul(rows)?.checked_mul(4)?;
    if rows != 0 && pitch < row_bytes { return None; }
    let span = if rows == 0 { 0 } else {
        (rows - 1).checked_mul(pitch)?.checked_add(row_bytes)?
    };
    if span > input.len() { return None; }
    let mut output = Vec::new();
    output.try_reserve_exact(length).ok()?;
    if row_bytes != 0 {
        for row in 0..rows {
            let start = row * pitch;
            let pixels = input[start..start + row_bytes].chunks_exact(3);
            if bgr {
                output.extend(pixels.flat_map(|pixel| [pixel[2], pixel[1], pixel[0], 255]));
            } else {
                output.extend(pixels.flat_map(|pixel| [pixel[0], pixel[1], pixel[2], 255]));
            }
        }
    }
    Some(output)
}

/// Expand a complete palette row at exact 2x horizontal scale.
pub(crate) fn double_palette_row(input: &[u8], output: &mut [u8], colors: &[[u8; 4]]) -> Option<()> {
    let length = input.len().checked_mul(8)?;
    if colors.len() < 256 || output.len() < length { return None; }
    for (&index, pair) in input.iter().zip(output[..length].chunks_exact_mut(8)) {
        let color = &colors[usize::from(index)];
        pair[..4].copy_from_slice(color);
        pair[4..].copy_from_slice(color);
    }
    Some(())
}

/// Scale opaque four-byte pixels at exact 2x horizontal scale without allocation.
pub(crate) fn scale_double_rgba_bytes(
    input: &[u8], output: &mut [u8], width: usize, rows: usize, pitch: usize,
    output_rows: usize, output_pitch: usize,
) -> Option<()> {
    if width == 0 || rows == 0 || output_rows == 0 { return None; }
    let source_bytes = width.checked_mul(4)?;
    let output_bytes = width.checked_mul(8)?;
    if pitch < source_bytes || output_pitch < output_bytes { return None; }
    let source_span = (rows - 1).checked_mul(pitch)?.checked_add(source_bytes)?;
    let output_span = (output_rows - 1).checked_mul(output_pitch)?.checked_add(output_bytes)?;
    if input.len() < source_span || output.len() < output_span { return None; }
    let (ys, yr) = (rows / output_rows, rows % output_rows);
    let yt = output_rows - yr;
    let (mut source_y, mut fraction_y) = (0, 0);
    let mut previous_y = None;
    for y in 0..output_rows {
        if previous_y == Some(source_y) {
            let previous = (y - 1) * output_pitch;
            output.copy_within(previous..previous + output_bytes, y * output_pitch);
        } else {
            let source = &input[source_y * pitch..source_y * pitch + source_bytes];
            let destination = &mut output[y * output_pitch..y * output_pitch + output_bytes];
            for (pixel, pair) in source.chunks_exact(4).zip(destination.chunks_exact_mut(8)) {
                pair[..4].copy_from_slice(pixel);
                pair[4..].copy_from_slice(pixel);
            }
        }
        previous_y = Some(source_y);
        source_y += ys;
        if fraction_y >= yt { fraction_y -= yt; source_y += 1; }
        else { fraction_y += yr; }
    }
    Some(())
}

#[cfg(test)]
mod tests {
    use super::swap_bgra_bytes;

    #[test]
    fn all_byte_values_at_word_boundaries_match_independent_pixel_oracle() {
        for length in 0..=71 {
            for seed in 0..=255u16 {
                let before: Vec<u8> = (0..length)
                    .map(|index| (index * 73 + usize::from(seed)) as u8).collect();
                let mut expected = before.clone();
                for start in (0..length / 4).map(|pixel| pixel * 4) {
                    expected[start] = before[start + 2];
                    expected[start + 2] = before[start];
                }
                let mut actual = before.clone();
                swap_bgra_bytes(&mut actual);
                assert_eq!(actual, expected, "length={length} seed={seed}");
                swap_bgra_bytes(&mut actual);
                assert_eq!(actual, before);
            }
        }
    }

    #[test]
    fn unaligned_subslices_preserve_guard_bytes_and_partial_pixels() {
        for offset in 0..16 {
            for length in 0..=71 {
                for seed in [0usize, 1, 127, 128, 255] {
                    let before: Vec<u8> = (0..offset + length + 19)
                        .map(|index| (index * 53 + seed) as u8).collect();
                    let mut expected = before.clone();
                    for pixel in 0..length / 4 {
                        let start = offset + pixel * 4;
                        expected[start] = before[start + 2];
                        expected[start + 2] = before[start];
                    }
                    let mut actual = before.clone();
                    swap_bgra_bytes(&mut actual[offset..offset + length]);
                    assert_eq!(actual, expected,
                        "offset={offset} length={length} seed={seed}");
                    swap_bgra_bytes(&mut actual[offset..offset + length]);
                    assert_eq!(actual, before);
                }
            }
        }
    }

    #[test]
    fn rgb24_rows_match_independent_channel_oracle() {
        for width in 0..18usize {
            for rows in 0..9usize {
                for padding in [0, 1, 13, 129] {
                    let pitch = width * 3 + padding;
                    let span = if rows == 0 { 0 } else { (rows - 1) * pitch + width * 3 };
                    let input: Vec<u8> = (0..span).map(|index| (index * 37 + 91) as u8).collect();
                    for bgr in [false, true] {
                        let mut expected = Vec::new();
                        for row in 0..rows {
                            for x in 0..width {
                                let start = row * pitch + x * 3;
                                let red = input[start + if bgr { 2 } else { 0 }];
                                let blue = input[start + if bgr { 0 } else { 2 }];
                                expected.extend_from_slice(&[red, input[start + 1], blue, 255]);
                            }
                        }
                        assert_eq!(super::expand_rgb24_rows(&input, width, rows, pitch, bgr), Some(expected));
                    }
                }
            }
        }
    }

    #[test]
    fn rgb24_rows_reject_invalid_bounds_and_overflow() {
        let input = [0xa5; 23];
        for (width, rows, pitch) in [(4, 2, 11), (4, 2, 24), (usize::MAX, 1, usize::MAX),
            (1, usize::MAX, 3), (1, 3, usize::MAX)] {
            assert_eq!(super::expand_rgb24_rows(&input, width, rows, pitch, false), None);
            assert_eq!(input, [0xa5; 23]);
        }
    }

    #[test]
    fn double_palette_row_preserves_colors_alpha_and_tail() {
        let colors: Vec<[u8; 4]> = (0..256usize)
            .map(|index| [index as u8, (index * 17) as u8, (index * 37) as u8, (index * 53) as u8]).collect();
        for width in 0..=513usize {
            let input: Vec<u8> = (0..width).map(|index| (index * 73 + 91) as u8).collect();
            let mut expected = Vec::new();
            for &index in &input {
                expected.extend_from_slice(&colors[usize::from(index)]);
                expected.extend_from_slice(&colors[usize::from(index)]);
            }
            expected.extend_from_slice(&[0xa5; 19]);
            let mut actual = vec![0xa5; width * 8 + 19];
            assert_eq!(super::double_palette_row(&input, &mut actual, &colors), Some(()));
            assert_eq!(actual, expected);
        }
        let mut output = [0xa5; 15];
        assert_eq!(super::double_palette_row(&[0, 1], &mut output, &colors), None);
        assert_eq!(super::double_palette_row(&[0], &mut output, &colors[..255]), None);
        assert_eq!(output, [0xa5; 15]);
    }

    #[test]
    fn truecolor_double_scaling_matches_independent_mapping_and_guards() {
        for width in 1..18usize {
            for rows in 1..10usize {
                for output_rows in 1..20usize {
                    for padding in [0, 13] {
                        let pitch = width * 4 + padding;
                        let output_pitch = width * 8 + padding;
                        let input: Vec<u8> = (0..(rows - 1) * pitch + width * 4)
                            .map(|index| (index * 73 + 91) as u8).collect();
                        let mut expected = vec![0xa5; (output_rows - 1) * output_pitch + width * 8 + 19];
                        for y in 0..output_rows {
                            for x in 0..width * 2 {
                                let source = (y * rows / output_rows) * pitch + (x / 2) * 4;
                                let destination = y * output_pitch + x * 4;
                                expected[destination..destination + 4].copy_from_slice(&input[source..source + 4]);
                            }
                        }
                        let mut actual = vec![0xa5; expected.len()];
                        assert_eq!(super::scale_double_rgba_bytes(&input, &mut actual, width, rows, pitch,
                            output_rows, output_pitch), Some(()));
                        assert_eq!(actual, expected);
                    }
                }
            }
        }
    }

    #[test]
    fn truecolor_double_scaling_rejects_all_ranges_before_writes() {
        for (width, rows, pitch, output_rows, output_pitch) in [
            (2, 2, 7, 2, 16), (2, 2, 8, 2, 15), (2, 2, 8, 2, 16),
            (usize::MAX, 1, usize::MAX, 1, usize::MAX),
            (1, 3, usize::MAX, 3, 8), (1, 1, 4, usize::MAX, 8),
            (0, 1, 0, 1, 0), (1, 0, 4, 1, 8), (1, 1, 4, 0, 8),
        ] {
            let mut output = [0xa5; 31];
            assert_eq!(super::scale_double_rgba_bytes(&[0; 16], &mut output, width, rows, pitch,
                output_rows, output_pitch), None);
            assert_eq!(output, [0xa5; 31]);
        }
    }
}

/// Integer nearest-neighbor expansion of opaque four-byte pixels.
/// All source and destination spans are checked before the first write.
pub(crate) fn scale_integer_rgba_bytes<const FACTOR: usize>(
    input: &[u8],
    output: &mut [u8],
    width: usize,
    rows: usize,
    pitch: usize,
    output_rows: usize,
    output_pitch: usize,
) -> Option<()> {
    if !(3..=4).contains(&FACTOR) || width == 0 || rows == 0 || output_rows == 0 {
        return None;
    }
    let source_bytes = width.checked_mul(4)?;
    let destination_bytes = source_bytes.checked_mul(FACTOR)?;
    if pitch < source_bytes || output_pitch < destination_bytes {
        return None;
    }
    let source_end = (rows - 1).checked_mul(pitch)?.checked_add(source_bytes)?;
    let destination_end = (output_rows - 1)
        .checked_mul(output_pitch)?
        .checked_add(destination_bytes)?;
    if source_end > input.len() || destination_end > output.len() {
        return None;
    }
    let mut previous = None;
    // The 3x path uses the same exact floor-mapping recurrence as the
    // double-scaling kernel. Subtraction before addition keeps the fraction
    // bounded without overflow. Keep 4x's measured baseline unchanged.
    let (row_step, row_remainder) = if FACTOR == 3 {
        (rows / output_rows, rows % output_rows)
    } else { (0, 0) };
    let row_threshold = output_rows - row_remainder;
    let (mut stepped_row, mut row_fraction) = (0usize, 0usize);
    for y in 0..output_rows {
        let source_y = if FACTOR == 3 { stepped_row } else {
            ((y as u128 * rows as u128) / output_rows as u128) as usize
        };
        let destination_start = y * output_pitch;
        if previous == Some(source_y) {
            let previous_start = destination_start - output_pitch;
            output.copy_within(
                previous_start..previous_start + destination_bytes,
                destination_start,
            );
        } else {
            let source_start = source_y * pitch;
            let source = &input[source_start..source_start + source_bytes];
            let destination =
                &mut output[destination_start..destination_start + destination_bytes];
            for (pixel, expanded) in source
                .chunks_exact(4)
                .zip(destination.chunks_exact_mut(4 * FACTOR))
            {
                for copy in expanded.chunks_exact_mut(4) {
                    copy.copy_from_slice(pixel);
                }
            }
        }
        previous = Some(source_y);
        if FACTOR == 3 {
            stepped_row += row_step;
            if row_fraction >= row_threshold {
                row_fraction -= row_threshold;
                stepped_row += 1;
            } else {
                row_fraction += row_remainder;
            }
        }
    }
    Some(())
}

/// Copy already independent pixel snapshots, preserving all row padding.
/// Matching complete byte-channel formats require no decode/encode round trip.
pub(crate) fn copy_snapshot_byte_rows(
    input: &[u8],
    output: &mut [u8],
    row_bytes: usize,
    rows: usize,
    source_pitch: usize,
    destination_pitch: usize,
) -> Option<()> {
    if row_bytes == 0 || rows == 0 {
        return Some(());
    }
    if source_pitch < row_bytes || destination_pitch < row_bytes {
        return None;
    }
    let source_span = (rows - 1).checked_mul(source_pitch)?.checked_add(row_bytes)?;
    let destination_span = (rows - 1).checked_mul(destination_pitch)?.checked_add(row_bytes)?;
    if source_span > input.len() || destination_span > output.len() {
        return None;
    }
    if source_pitch == row_bytes && destination_pitch == row_bytes {
        output[..destination_span].copy_from_slice(&input[..source_span]);
    } else {
        for row in 0..rows {
            let source = row * source_pitch;
            let destination = row * destination_pitch;
            output[destination..destination + row_bytes]
                .copy_from_slice(&input[source..source + row_bytes]);
        }
    }
    Some(())
}

#[cfg(test)]
mod mapping_and_integer_scale_tests {
    use super::{byte_channel_mapping, scale_integer_rgba_bytes};

    #[test]
    fn channel_defaults_and_permutations_match_packed_mask_oracle() {
        let mut permutations = Vec::new();
        for a in 0..4 {
            for b in 0..4 {
                for c in 0..4 {
                    for d in 0..4 {
                        if a != b && a != c && a != d && b != c && b != d && c != d {
                            permutations.push([a, b, c, d]);
                        }
                    }
                }
            }
        }
        for src_order in &permutations {
            for dst_order in &permutations {
                for src_present in 0..16 {
                    for dst_present in 0..16 {
                        let source = std::array::from_fn(|i| {
                            if src_present & (1 << i) != 0 { 255u32 << (src_order[i] * 8) } else { 0 }
                        });
                        let destination = std::array::from_fn(|i| {
                            if dst_present & (1 << i) != 0 { 255u32 << (dst_order[i] * 8) } else { 0 }
                        });
                        let (mapping, defaults) = byte_channel_mapping(source, destination);
                        for pixel in [[0, 1, 127, 255], [255, 128, 17, 0], [1, 254, 255, 128]] {
                            let packed = u32::from_le_bytes(pixel);
                            let mut expected = 0u32;
                            for component in 0..4 {
                                if destination[component] == 0 { continue; }
                                let value = if source[component] == 0 {
                                    if component == 3 { 255 } else { 0 }
                                } else {
                                    (packed & source[component]) >> source[component].trailing_zeros()
                                };
                                expected |= value << destination[component].trailing_zeros();
                            }
                            let actual: [u8; 4] = std::array::from_fn(|byte| {
                                mapping[byte].map_or(defaults[byte], |index| pixel[index])
                            });
                            assert_eq!(actual, expected.to_le_bytes());
                        }
                    }
                }
            }
        }
    }

    fn check_scale<const FACTOR: usize>() {
        for width in [1, 2, 7, 15, 16, 17] {
            for rows in 1..20usize {
                for output_rows in 1..41usize {
                    for padding in [0, 1, 13] {
                        let pitch = width * 4 + padding;
                        let output_pitch = width * 4 * FACTOR + padding;
                        let source: Vec<u8> = (0..(rows - 1) * pitch + width * 4)
                            .map(|i| (i * 73 + 91) as u8).collect();
                        let mut expected = vec![0xa5; output_rows * output_pitch + 19];
                        for y in 0..output_rows {
                            for x in 0..width * FACTOR {
                                let src = (y * rows / output_rows) * pitch + (x / FACTOR) * 4;
                                let dst = y * output_pitch + x * 4;
                                expected[dst..dst + 4].copy_from_slice(&source[src..src + 4]);
                            }
                        }
                        let mut actual = vec![0xa5; expected.len()];
                        assert_eq!(scale_integer_rgba_bytes::<FACTOR>(
                            &source, &mut actual, width, rows, pitch, output_rows, output_pitch,
                        ), Some(()));
                        assert_eq!(actual, expected);
                    }
                }
            }
        }
        for (width, rows, pitch, output_rows, output_pitch) in [
            (0, 1, 0, 1, 0), (1, 0, 4, 1, 4 * FACTOR),
            (1, 1, 4, 0, 4 * FACTOR), (2, 2, 7, 2, 8 * FACTOR),
            (1, 2, 4, 2, 4 * FACTOR - 1),
            (usize::MAX, 1, usize::MAX, 1, usize::MAX),
            (1, usize::MAX, 4, 1, 4 * FACTOR),
            (1, 1, 4, usize::MAX, 4 * FACTOR),
        ] {
            let mut output = [0xa5; 63];
            assert_eq!(scale_integer_rgba_bytes::<FACTOR>(
                &[0; 8], &mut output, width, rows, pitch, output_rows, output_pitch,
            ), None);
            assert_eq!(output, [0xa5; 63]);
        }
    }

    #[test]
    fn integer_scaling_matches_floor_mapping_padding_and_bounds() {
        check_scale::<3>();
        check_scale::<4>();
    }
}
