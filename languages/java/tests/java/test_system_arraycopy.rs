use crate::helpers::run_main;

#[test]
fn system_arraycopy_copies_prefix_into_empty_destination() {
    let out = run_main(
        "int[] src = {10, 20, 30}; int[] dest = new int[3]; System.arraycopy(src, 0, dest, 0, 3); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn system_arraycopy_copies_middle_slice_to_start_of_destination() {
    let out = run_main(
        "int[] src = {1, 2, 3, 4, 5}; int[] dest = new int[3]; System.arraycopy(src, 1, dest, 0, 3); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn system_arraycopy_copies_tail_segment_into_destination() {
    let out = run_main(
        "int[] src = {5, 6, 7, 8, 9}; int[] dest = new int[2]; System.arraycopy(src, 3, dest, 0, 2); System.out.println(dest[0]); System.out.println(dest[1]);",
    );
    assert_eq!(out, vec!["8", "9"]);
}

#[test]
fn system_arraycopy_zero_length_leaves_destination_untouched() {
    let out = run_main(
        "int[] src = {1, 2, 3}; int[] dest = {9, 9, 9}; System.arraycopy(src, 0, dest, 0, 0); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["9", "9"]);
}

#[test]
fn system_arraycopy_single_element_copy() {
    let out = run_main(
        "int[] src = {42}; int[] dest = new int[1]; System.arraycopy(src, 0, dest, 0, 1); System.out.println(dest[0]);",
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn system_arraycopy_copies_into_middle_of_larger_destination() {
    let out = run_main(
        "int[] src = {11, 22}; int[] dest = {0, 0, 0, 0}; System.arraycopy(src, 0, dest, 1, 2); System.out.println(dest[1]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["11", "22"]);
}

#[test]
fn system_arraycopy_source_and_destination_are_same_array_non_overlapping() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5}; System.arraycopy(data, 0, data, 3, 2); System.out.println(data[3]); System.out.println(data[4]);",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn system_arraycopy_forward_overlap_shifts_elements_right() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5}; System.arraycopy(data, 0, data, 2, 3); System.out.println(data[2]); System.out.println(data[4]);",
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn system_arraycopy_forward_overlap_preserves_leading_elements() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5}; System.arraycopy(data, 0, data, 2, 3); System.out.println(data[0]); System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn system_arraycopy_backward_overlap_shifts_elements_left() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5}; System.arraycopy(data, 2, data, 0, 3); System.out.println(data[0]); System.out.println(data[2]);",
    );
    assert_eq!(out, vec!["3", "5"]);
}

#[test]
fn system_arraycopy_backward_overlap_preserves_trailing_elements() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5}; System.arraycopy(data, 2, data, 0, 3); System.out.println(data[3]); System.out.println(data[4]);",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn system_arraycopy_self_copy_zero_offset_is_identity() {
    let out = run_main(
        "int[] data = {7, 8, 9}; System.arraycopy(data, 0, data, 0, 3); System.out.println(data[0]); System.out.println(data[2]);",
    );
    assert_eq!(out, vec!["7", "9"]);
}

#[test]
fn system_arraycopy_copies_negative_source_values() {
    let out = run_main(
        "int[] src = {-3, -2, -1}; int[] dest = new int[3]; System.arraycopy(src, 0, dest, 0, 3); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["-3", "-1"]);
}

#[test]
fn system_arraycopy_copies_mixed_sign_values() {
    let out = run_main(
        "int[] src = {-1, 0, 1}; int[] dest = new int[3]; System.arraycopy(src, 0, dest, 0, 3); System.out.println(dest[1]);",
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn system_arraycopy_partial_copy_does_not_touch_remaining_destination_slots() {
    let out = run_main(
        "int[] src = {100, 200}; int[] dest = {1, 2, 3, 4}; System.arraycopy(src, 0, dest, 0, 2); System.out.println(dest[2]); System.out.println(dest[3]);",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn system_arraycopy_from_second_source_index() {
    let out = run_main(
        "int[] src = {9, 8, 7, 6}; int[] dest = new int[2]; System.arraycopy(src, 2, dest, 0, 2); System.out.println(dest[0]); System.out.println(dest[1]);",
    );
    assert_eq!(out, vec!["7", "6"]);
}

#[test]
fn system_arraycopy_into_second_destination_index() {
    let out = run_main(
        "int[] src = {4, 5}; int[] dest = {0, 0, 0, 0}; System.arraycopy(src, 0, dest, 2, 2); System.out.println(dest[2]); System.out.println(dest[3]);",
    );
    assert_eq!(out, vec!["4", "5"]);
}

#[test]
fn system_arraycopy_longer_source_slice_of_five_elements() {
    let out = run_main(
        "int[] src = {2, 4, 6, 8, 10}; int[] dest = new int[5]; System.arraycopy(src, 0, dest, 0, 5); System.out.println(dest[0]); System.out.println(dest[4]);",
    );
    assert_eq!(out, vec!["2", "10"]);
}

#[test]
fn system_arraycopy_overwrites_destination_completely() {
    let out = run_main(
        "int[] src = {99, 88, 77}; int[] dest = {1, 2, 3}; System.arraycopy(src, 0, dest, 0, 3); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["99", "77"]);
}

#[test]
fn system_arraycopy_forward_overlap_single_step() {
    let out = run_main(
        "int[] data = {1, 2, 3}; System.arraycopy(data, 0, data, 1, 2); System.out.println(data[1]); System.out.println(data[2]);",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn system_arraycopy_backward_overlap_single_step() {
    let out = run_main(
        "int[] data = {1, 2, 3}; System.arraycopy(data, 1, data, 0, 2); System.out.println(data[0]); System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn system_arraycopy_copies_zeros_from_source() {
    let out = run_main(
        "int[] src = {0, 0, 0}; int[] dest = {5, 5, 5}; System.arraycopy(src, 0, dest, 0, 3); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn system_arraycopy_copies_max_int_value() {
    let out = run_main(
        "int[] src = {2147483647}; int[] dest = new int[1]; System.arraycopy(src, 0, dest, 0, 1); System.out.println(dest[0]);",
    );
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn system_arraycopy_copies_min_int_value() {
    let out = run_main(
        "int[] src = {-2147483648}; int[] dest = new int[1]; System.arraycopy(src, 0, dest, 0, 1); System.out.println(dest[0]);",
    );
    assert_eq!(out, vec!["-2147483648"]);
}

#[test]
fn system_arraycopy_two_element_overlap_forward_in_four_slot_array() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4}; System.arraycopy(data, 0, data, 1, 2); System.out.println(data[1]); System.out.println(data[2]);",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn system_arraycopy_two_element_overlap_backward_in_four_slot_array() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4}; System.arraycopy(data, 2, data, 0, 2); System.out.println(data[0]); System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn system_arraycopy_three_element_copy_from_offset_one() {
    let out = run_main(
        "int[] src = {0, 10, 20, 30, 40}; int[] dest = new int[3]; System.arraycopy(src, 1, dest, 0, 3); System.out.println(dest[0]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn system_arraycopy_replaces_interior_segment_only() {
    let out = run_main(
        "int[] src = {50, 60}; int[] dest = {1, 2, 3, 4}; System.arraycopy(src, 0, dest, 1, 2); System.out.println(dest[0]); System.out.println(dest[1]); System.out.println(dest[2]);",
    );
    assert_eq!(out, vec!["1", "50", "60"]);
}

#[test]
fn system_arraycopy_forward_overlap_in_six_element_buffer() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5, 6}; System.arraycopy(data, 1, data, 3, 3); System.out.println(data[3]); System.out.println(data[5]);",
    );
    assert_eq!(out, vec!["2", "4"]);
}

#[test]
fn system_arraycopy_backward_overlap_in_six_element_buffer() {
    // JLS: overlapping copy behaves as if through a temporary —
    // src[3..6]=[4,5,6] lands at dest[1..4], so data == {1,4,5,6,5,6}.
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5, 6}; System.arraycopy(data, 3, data, 1, 3); System.out.println(data[1]); System.out.println(data[3]);",
    );
    assert_eq!(out, vec!["4", "6"]);
}

#[test]
fn system_arraycopy_copies_alternating_sign_pattern() {
    let out = run_main(
        "int[] src = {1, -1, 1, -1}; int[] dest = new int[4]; System.arraycopy(src, 0, dest, 0, 4); System.out.println(dest[0]); System.out.println(dest[3]);",
    );
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn system_arraycopy_leaves_prefix_when_copying_to_higher_index() {
    let out = run_main(
        "int[] data = {9, 8, 7, 6}; System.arraycopy(data, 0, data, 2, 2); System.out.println(data[0]); System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["9", "8"]);
}

#[test]
fn system_arraycopy_copies_identical_source_and_dest_lengths_of_one() {
    let out = run_main(
        "int[] src = {123}; int[] dest = {0}; System.arraycopy(src, 0, dest, 0, 1); System.out.println(dest[0]);",
    );
    assert_eq!(out, vec!["123"]);
}

#[test]
fn system_arraycopy_shift_right_by_one_in_three_element_array() {
    let out = run_main(
        "int[] data = {10, 20, 30}; System.arraycopy(data, 0, data, 1, 2); System.out.println(data[1]); System.out.println(data[2]);",
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn system_arraycopy_shift_left_by_one_in_three_element_array() {
    let out = run_main(
        "int[] data = {10, 20, 30}; System.arraycopy(data, 1, data, 0, 2); System.out.println(data[0]); System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["20", "30"]);
}

#[test]
fn system_arraycopy_copies_second_half_to_first_half() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4}; System.arraycopy(data, 2, data, 0, 2); System.out.println(data[0]); System.out.println(data[1]);",
    );
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn system_arraycopy_copies_first_half_to_second_half_non_overlap() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4}; System.arraycopy(data, 0, data, 2, 2); System.out.println(data[2]); System.out.println(data[3]);",
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn system_arraycopy_repeated_values_all_copied() {
    let out = run_main(
        "int[] src = {5, 5, 5, 5}; int[] dest = new int[4]; System.arraycopy(src, 0, dest, 0, 4); System.out.println(dest[0]); System.out.println(dest[3]);",
    );
    assert_eq!(out, vec!["5", "5"]);
}

#[test]
fn system_arraycopy_large_values_near_int_max() {
    let out = run_main(
        "int[] src = {2000000000, 2000000001}; int[] dest = new int[2]; System.arraycopy(src, 0, dest, 0, 2); System.out.println(dest[0]); System.out.println(dest[1]);",
    );
    assert_eq!(out, vec!["2000000000", "2000000001"]);
}

#[test]
fn system_arraycopy_forward_overlap_three_to_end_in_five_slots() {
    let out = run_main(
        "int[] data = {1, 2, 3, 4, 5}; System.arraycopy(data, 0, data, 2, 3); System.out.println(data[2]); System.out.println(data[3]); System.out.println(data[4]);",
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}
