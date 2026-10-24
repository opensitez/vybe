//! Display-name aggregation for explicitly requested emission diagnostics.
//! A name is a reporting label, never a helper/linker identity certificate.

use std::collections::BTreeMap;
use std::fmt::Write;

use crate::Chunk;

#[derive(Default)]
struct Group {
    chunks: usize,
    code_bytes: usize,
    capacity_bytes: usize,
    smallest_body_bytes: usize,
    largest_body_bytes: usize,
}

pub(super) fn write_groups(output: &mut String, chunks: &[Chunk]) {
    let mut groups: BTreeMap<&str, Group> = BTreeMap::new();
    for chunk in chunks {
        let group = groups.entry(chunk.name.as_str()).or_default();
        let bytes = chunk.code.len();
        group.smallest_body_bytes = if group.chunks == 0 {
            bytes
        } else {
            group.smallest_body_bytes.min(bytes)
        };
        group.largest_body_bytes = group.largest_body_bytes.max(bytes);
        group.chunks += 1;
        group.code_bytes = group.code_bytes.saturating_add(bytes);
        group.capacity_bytes = group.capacity_bytes.saturating_add(chunk.code.capacity());
    }
    let _ = writeln!(output,
        "emission_groups grouping=display_name_only not_linker_identity=true unique_names={} chunks={}",
        groups.len(), chunks.len());
    let mut ranked: Vec<_> = groups.into_iter().collect();
    ranked.sort_unstable_by(|left, right| {
        right.1.code_bytes.cmp(&left.1.code_bytes).then_with(|| left.0.cmp(right.0))
    });
    for (name, group) in ranked.into_iter().take(30) {
        let _ = writeln!(output,
            "emission_group name={name:?} chunks={} code_bytes={} code_capacity_bytes={} smallest_body_bytes={} largest_body_bytes={}",
            group.chunks, group.code_bytes, group.capacity_bytes,
            group.smallest_body_bytes, group.largest_body_bytes);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn groups_distinct_bodies_without_claiming_identity_or_mutating_code() {
        let mut first = Chunk::new("repeated");
        first.code.extend_from_slice(&[1, 2, 3]);
        let mut second = Chunk::new("repeated");
        second.code.extend_from_slice(&[4, 5]);
        let mut larger = Chunk::new("larger");
        larger.code.extend_from_slice(&[6; 8]);
        let chunks = vec![first, second, larger];
        let pointers: Vec<_> = chunks.iter().map(|chunk| chunk.code.as_ptr()).collect();
        let capacities: Vec<_> = chunks.iter().map(|chunk| chunk.code.capacity()).collect();
        let mut output = String::new();
        write_groups(&mut output, &chunks);
        assert!(output.contains("not_linker_identity=true unique_names=2 chunks=3"));
        let larger = output.find("name=\"larger\" chunks=1 code_bytes=8").unwrap();
        let repeated = output.find("name=\"repeated\" chunks=2 code_bytes=5").unwrap();
        assert!(larger < repeated);
        assert!(output.contains("smallest_body_bytes=2 largest_body_bytes=3"));
        for (index, chunk) in chunks.iter().enumerate() {
            assert_eq!(chunk.code.as_ptr(), pointers[index]);
            assert_eq!(chunk.code.capacity(), capacities[index]);
        }
        assert_eq!(chunks[0].code, [1, 2, 3]);
        assert_eq!(chunks[1].code, [4, 5]);
    }

    #[test]
    fn empty_input_has_only_a_header() {
        let mut output = String::new();
        write_groups(&mut output, &[]);
        assert_eq!(output.lines().count(), 1);
        assert!(output.contains("unique_names=0 chunks=0"));
    }

    #[test]
    fn equal_sizes_use_name_order_and_output_is_bounded() {
        let mut chunks = Vec::new();
        for index in (0..40).rev() {
            let mut chunk = Chunk::new(format!("group{index:02}"));
            chunk.code.push(0);
            chunks.push(chunk);
        }
        let mut output = String::new();
        write_groups(&mut output, &chunks);
        assert_eq!(output.lines().count(), 31);
        assert!(output.lines().nth(1).unwrap().contains("group00"));
        assert!(output.lines().last().unwrap().contains("group29"));
    }
}
