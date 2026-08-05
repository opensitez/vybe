//! Kotlin's `common:kotlin.*` emit targets.
//!
//! The shared dispatcher hands a `common:<name>` target it does not recognise
//! to the language that declared it, which is how a language owns an emitter
//! without shared code learning its name.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    if name.starts_with("java.") {
        return vybe_language_java::emitter::dispatch::dispatch(name, chunks, current, argc, line);
    }
    match name {
        "kotlin.sorted_set_of" => { crate::emitter::hof::emit_sorted_set_of(chunks, current, argc, line); true }
        "kotlin.sorted_map_of" => { crate::emitter::hof::emit_sorted_map_of(chunks, current, argc, line); true }
        "kotlin.kt_filter" => { crate::emitter::hof::emit_kt_filter(chunks, current, argc, line); true }
        "kotlin.kt_filter_not" => { crate::emitter::hof::emit_kt_filter_not(chunks, current, argc, line); true }
        "kotlin.kt_map_hof" => { crate::emitter::hof::emit_kt_map_hof(chunks, current, argc, line); true }
        "kotlin.kt_for_each" => { crate::emitter::hof::emit_kt_for_each(chunks, current, argc, line); true }
        "kotlin.dict_set_tracked" => { crate::emitter::maps::emit_dict_set_stmt(chunks, current, argc, line); true }
        "kotlin.map_for_each" => { crate::emitter::maps::emit_map_for_each(chunks, current, argc, line); true }
        "kotlin.null_array" => { crate::emitter::hof::emit_null_array(chunks, current, argc, line); true }
        "kotlin.array_init" => { crate::emitter::hof::emit_array_init(chunks, current, argc, line); true }
        "kotlin.fill" => { crate::emitter::hof::emit_fill(chunks, current, argc, line); true }
        "kotlin.remove_if" => { crate::emitter::hof::emit_remove_if(chunks, current, argc, line); true }
        "kotlin.retain_if" => { crate::emitter::hof::emit_retain_if(chunks, current, argc, line); true }
        "kotlin.map_replace" => { crate::emitter::maps::emit_map_replace(chunks, current, argc, line); true }
        "kotlin.map_indexed_not_null" => { crate::emitter::hof::emit_map_indexed_not_null(chunks, current, argc, line); true }
        "kotlin.size_any" => { crate::emitter::hof::emit_size_any(chunks, current, line); true }
        "kotlin.dict_delete_full" => { crate::emitter::maps::emit_dict_delete_full(chunks, current, argc, line); true }
        "kotlin.no_such_element_exception" => {
            crate::emitter::nullability::emit_exception(chunks, current, argc, "NoSuchElementException", line);
            true
        }
        "kotlin.unsupported_operation_exception" => {
            crate::emitter::nullability::emit_exception(chunks, current, argc, "UnsupportedOperationException", line);
            true
        }
        "kotlin.reduce_throwing" => { crate::emitter::hof::emit_reduce_throwing(chunks, current, argc, line); true }
        "kotlin.is_blank" => { crate::emitter::strings::emit_is_blank(chunks, current, argc, line); true }
        "kotlin.is_not_blank" => { crate::emitter::strings::emit_is_not_blank(chunks, current, argc, line); true }
        "kotlin.is_null_or_empty" => { crate::emitter::strings::emit_is_null_or_empty(chunks, current, argc, line); true }
        "kotlin.is_null_or_blank" => { crate::emitter::strings::emit_is_null_or_blank(chunks, current, argc, line); true }
        "kotlin.remove_prefix" => { crate::emitter::strings::emit_remove_prefix(chunks, current, argc, line); true }
        "kotlin.remove_suffix" => { crate::emitter::strings::emit_remove_suffix(chunks, current, argc, line); true }
        "kotlin.substring_before" => { crate::emitter::strings::emit_substring_around(chunks, current, argc, false, false, line); true }
        "kotlin.substring_after" => { crate::emitter::strings::emit_substring_around(chunks, current, argc, true, false, line); true }
        "kotlin.substring_before_last" => { crate::emitter::strings::emit_substring_around(chunks, current, argc, false, true, line); true }
        "kotlin.substring_after_last" => { crate::emitter::strings::emit_substring_around(chunks, current, argc, true, true, line); true }
        "kotlin.lines" => { crate::emitter::strings::emit_lines(chunks, current, argc, line); true }
        "kotlin.replace_range" => { crate::emitter::strings::emit_replace_range(chunks, current, argc, line); true }
        "kotlin.compare_to" => { crate::emitter::strings::emit_compare_to(chunks, current, argc, line); true }
        "kotlin.equals_ic" => { crate::emitter::strings::emit_ignore_case_op(chunks, current, argc, "equals", line); true }
        "kotlin.contains_ic" => { crate::emitter::strings::emit_ignore_case_op(chunks, current, argc, "includes", line); true }
        "kotlin.starts_with_ic" => { crate::emitter::strings::emit_ignore_case_op(chunks, current, argc, "startsWith", line); true }
        "kotlin.ends_with_ic" => { crate::emitter::strings::emit_ignore_case_op(chunks, current, argc, "endsWith", line); true }
        "kotlin.index_of_any" => { crate::emitter::strings::emit_index_of_any(chunks, current, argc, false, line); true }
        "kotlin.last_index_of_any" => { crate::emitter::strings::emit_index_of_any(chunks, current, argc, true, line); true }
        "kotlin.to_boolean" => { crate::emitter::strings::emit_to_boolean(chunks, current, argc, line); true }
        "kotlin.to_boolean_strict_or_null" => { crate::emitter::strings::emit_to_boolean_strict_or_null(chunks, current, argc, line); true }
        "kotlin.is_digit" => { crate::emitter::strings::emit_is_digit(chunks, current, argc, line); true }
        "kotlin.is_letter" => { crate::emitter::strings::emit_is_letter(chunks, current, argc, line); true }
        "kotlin.trim_indent" => { crate::emitter::strings::emit_trim_indent(chunks, current, argc, line); true }
        "kotlin.trim_margin" => { crate::emitter::strings::emit_trim_margin(chunks, current, argc, line); true }
        "kotlin.slice_any" => { crate::emitter::hof::emit_slice_any(chunks, current, argc, line); true }
        "kotlin.reversed_any" => { crate::emitter::strings::emit_reversed_any(chunks, current, argc, line); true }
        "kotlin.region_matches" => { crate::emitter::strings::emit_region_matches(chunks, current, argc, line); true }
        "kotlin.index_of_from" => { crate::emitter::strings::emit_index_of_from(chunks, current, argc, false, line); true }
        "kotlin.last_index_of_from" => { crate::emitter::strings::emit_index_of_from(chunks, current, argc, true, line); true }
        "kotlin.chunked" => { crate::emitter::strings::emit_chunked_windowed(chunks, current, argc, false, line); true }
        "kotlin.windowed" => { crate::emitter::strings::emit_chunked_windowed(chunks, current, argc, true, line); true }
        "kotlin.windowed_partial" => { crate::emitter::strings::emit_chunked_windowed_ex(chunks, current, argc, true, true, line); true }
        "kotlin.slice_range_value" => { crate::emitter::strings::emit_slice_range_value(chunks, current, argc, line); true }
        "kotlin.index_of_any_recv" => { crate::emitter::strings::emit_index_of_any_recv(chunks, current, argc, false, line); true }
        "kotlin.last_index_of_any_recv" => { crate::emitter::strings::emit_index_of_any_recv(chunks, current, argc, true, line); true }
        "kotlin.to_byte_or_null" => { crate::emitter::strings::emit_to_bounded_or_null(chunks, current, -128, 127, line); true }
        "kotlin.to_short_or_null" => { crate::emitter::strings::emit_to_bounded_or_null(chunks, current, -32768, 32767, line); true }
        "kotlin.take_last_while" => { crate::emitter::hof::emit_last_while_split(chunks, current, true, line); true }
        "kotlin.drop_last_while" => { crate::emitter::hof::emit_last_while_split(chunks, current, false, line); true }
        "kotlin.trim_any" => {
            if argc <= 1 {
                let idx = chunks[current].add_import("ecma:string", "trim");
                chunks[current].emit_call(idx, 1, line);
            } else {
                crate::emitter::strings::emit_trim_chars(chunks, current, argc, true, true, line);
            }
            true
        }
        "kotlin.trim_start_any" => {
            if argc <= 1 {
                let idx = chunks[current].add_import("ecma:string", "trimStart");
                chunks[current].emit_call(idx, 1, line);
            } else {
                crate::emitter::strings::emit_trim_chars(chunks, current, argc, true, false, line);
            }
            true
        }
        "kotlin.trim_end_any" => {
            if argc <= 1 {
                let idx = chunks[current].add_import("ecma:string", "trimEnd");
                chunks[current].emit_call(idx, 1, line);
            } else {
                crate::emitter::strings::emit_trim_chars(chunks, current, argc, false, true, line);
            }
            true
        }
        "kotlin.split_limit" => { crate::emitter::strings::emit_split_limit(chunks, current, argc, line); true }
        "kotlin.char_at_throwing" => { crate::emitter::strings::emit_char_at_throwing(chunks, current, argc, line); true }
        "kotlin.substring_throwing" => { crate::emitter::strings::emit_substring_throwing(chunks, current, argc, line); true }
        "kotlin.contains_any" => { crate::emitter::strings::emit_contains_any(chunks, current, argc, line); true }
        "kotlin.to_int_radix" => {
            if argc >= 2 {
                crate::emitter::strings::emit_parse_radix(chunks, current, argc, false, line);
            } else {
                // NOT a bare trunc: a string receiver must THROW
                // NumberFormatException on garbage (`"x".toInt()` was NaN,
                // so `runCatching` never caught anything).
                crate::emitter::numbers::emit_to_int_throwing(chunks, current, line);
            }
            true
        }
        "kotlin.to_int_or_null_radix" => {
            if argc >= 2 {
                crate::emitter::strings::emit_parse_radix(chunks, current, argc, true, line);
            } else {
                crate::emitter::strings::emit_strict_int_or_null(chunks, current, argc, line);
            }
            true
        }
        "kotlin.find" => { crate::emitter::hof::emit_find(chunks, current, argc, line); true }
        "kotlin.find_last_any" => { crate::emitter::hof::emit_find_last_any(chunks, current, argc, line); true }
        "kotlin.remove_any" => { crate::emitter::maps::emit_remove_any(chunks, current, argc, line); true }
        "kotlin.clear_any" => { crate::emitter::maps::emit_clear_any(chunks, current, argc, line); true }
        "kotlin.map_put" => { crate::emitter::maps::emit_map_put(chunks, current, argc, line); true }
        "kotlin.safe_get" => { crate::emitter::maps::emit_safe_get(chunks, current, argc, line); true }
        "kotlin.dict_get_null" => { crate::emitter::maps::emit_dict_get_null(chunks, current, argc, line); true }
        "kotlin.list_set_checked" => { crate::emitter::maps::emit_list_set_checked(chunks, current, argc, line); true }
        "kotlin.concat_to_string" => { crate::emitter::collections::emit_concat_to_string(chunks, current, argc, line); true }
        "kotlin.reverse_in_place" => { crate::emitter::collections::emit_reverse_in_place(chunks, current, argc, line); true }
        "kotlin.sorted_copy" => { crate::emitter::collections::emit_sorted_copy(chunks, current, argc, line); true }
        "kotlin.sorted_desc_copy" => { crate::emitter::collections::emit_sorted_desc_copy(chunks, current, argc, line); true }
        "kotlin.sort_range" => { crate::emitter::collections::emit_sort_range(chunks, current, argc, line); true }
        "kotlin.binary_search_range" => { crate::emitter::collections::emit_binary_search_range(chunks, current, argc, line); true }
        "kotlin.to_byte_array" => { crate::emitter::collections::emit_to_byte_array(chunks, current, argc, line); true }
        "kotlin.to_byte_wrap" => { crate::emitter::numbers::emit_wrap_int(chunks, current, 8, true, line); true }
        "kotlin.to_short_wrap" => { crate::emitter::numbers::emit_wrap_int(chunks, current, 16, true, line); true }
        "kotlin.to_ubyte" => { crate::emitter::numbers::emit_wrap_int(chunks, current, 8, false, line); true }
        "kotlin.to_ushort" => { crate::emitter::numbers::emit_wrap_int(chunks, current, 16, false, line); true }
        "kotlin.double_str" => { crate::emitter::numbers::emit_double_to_string(chunks, current, line); true }
        "kotlin.tuple_prop" => { crate::emitter::maps::emit_tuple_prop(chunks, current, argc, line); true }
        "kotlin.list_get_throwing" => { crate::emitter::maps::emit_list_get_throwing(chunks, current, argc, line); true }
        "kotlin.sub_list" => { crate::emitter::maps::emit_sub_list(chunks, current, argc, line); true }
        "kotlin.fold_right_indexed" => { crate::emitter::hof::emit_fold_right_indexed(chunks, current, argc, line); true }
        "kotlin.running_fold_indexed" => { crate::emitter::hof::emit_running_fold_indexed(chunks, current, argc, line); true }
        "kotlin.distinct" => { crate::emitter::hof::emit_distinct(chunks, current, argc, line); true }
        "kotlin.to_sorted_set" => { crate::emitter::hof::emit_to_sorted_set(chunks, current, argc, line); true }
        "kotlin.take_while" => { crate::emitter::hof::emit_take_while(chunks, current, argc, line); true }
        "kotlin.drop_while" => { crate::emitter::hof::emit_drop_while(chunks, current, argc, line); true }
        "kotlin.count" => { crate::emitter::hof::emit_count(chunks, current, argc, line); true }
        "kotlin.none" => { crate::emitter::hof::emit_none(chunks, current, argc, line); true }
        "kotlin.sum_of" => { crate::emitter::hof::emit_sum_of(chunks, current, argc, line); true }
        "kotlin.min_by_or_null" => { crate::emitter::hof::emit_min_by_or_null(chunks, current, argc, line); true }
        "kotlin.max_by_or_null" => { crate::emitter::hof::emit_max_by_or_null(chunks, current, argc, line); true }
        "kotlin.index_of_first" => { crate::emitter::hof::emit_index_of_first(chunks, current, argc, line); true }
        "kotlin.index_of_last" => { crate::emitter::hof::emit_index_of_last(chunks, current, argc, line); true }
        "kotlin.find_last" => { crate::emitter::hof::emit_find_last(chunks, current, argc, line); true }
        "kotlin.first" => { crate::emitter::hof::emit_first(chunks, current, argc, line); true }
        "kotlin.first_or_null" => { crate::emitter::hof::emit_first_or_null(chunks, current, argc, line); true }
        "kotlin.last" => { crate::emitter::hof::emit_last(chunks, current, argc, line); true }
        "kotlin.last_or_null" => { crate::emitter::hof::emit_last_or_null(chunks, current, argc, line); true }
        "kotlin.single" => { crate::emitter::hof::emit_single(chunks, current, argc, line); true }
        "kotlin.single_or_null" => { crate::emitter::hof::emit_single_or_null(chunks, current, argc, line); true }
        "kotlin.get_or_null" => { crate::emitter::hof::emit_get_or_null(chunks, current, argc, line); true }
        "kotlin.get_or_else" => { crate::emitter::hof::emit_get_or_else(chunks, current, argc, line); true }
        "kotlin.filter_not" => { crate::emitter::hof::emit_filter_not(chunks, current, argc, line); true }
        "kotlin.filter_indexed" => { crate::emitter::hof::emit_filter_indexed(chunks, current, argc, line); true }
        "kotlin.filter_not_null" => { crate::emitter::hof::emit_filter_not_null(chunks, current, argc, line); true }
        "kotlin.map_not_null" => { crate::emitter::hof::emit_map_not_null(chunks, current, argc, line); true }
        "kotlin.map_indexed" => { crate::emitter::hof::emit_map_indexed(chunks, current, argc, line); true }
        "kotlin.for_each_indexed" => { crate::emitter::hof::emit_for_each_indexed(chunks, current, argc, line); true }
        "kotlin.on_each" => { crate::emitter::hof::emit_on_each(chunks, current, argc, line); true }
        "kotlin.distinct_by" => { crate::emitter::hof::emit_distinct_by(chunks, current, argc, line); true }
        "kotlin.fold_right" => { crate::emitter::hof::emit_fold_right(chunks, current, argc, line); true }
        "kotlin.reduce_right" => { crate::emitter::hof::emit_reduce_right(chunks, current, argc, line); true }
        "kotlin.fold_indexed" => { crate::emitter::hof::emit_fold_indexed(chunks, current, argc, line); true }
        "kotlin.reduce_or_null" => { crate::emitter::hof::emit_reduce_or_null(chunks, current, argc, line); true }
        "kotlin.running_fold" => { crate::emitter::hof::emit_running_fold(chunks, current, argc, line); true }
        "kotlin.running_reduce" => { crate::emitter::hof::emit_running_reduce(chunks, current, argc, line); true }
        "kotlin.group_by" => { crate::emitter::hof::emit_group_by(chunks, current, argc, line); true }
        "kotlin.group_by_to" => { crate::emitter::hof::emit_group_by_to(chunks, current, argc, line); true }
        "kotlin.associate_by" => { crate::emitter::hof::emit_associate_by(chunks, current, argc, line); true }
        "kotlin.associate_by_to" => { crate::emitter::hof::emit_associate_by_to(chunks, current, argc, line); true }
        "kotlin.associate_with" => { crate::emitter::hof::emit_associate_with(chunks, current, argc, line); true }
        "kotlin.associate" => { crate::emitter::hof::emit_associate(chunks, current, argc, line); true }
        "kotlin.zip" => { crate::emitter::hof::emit_zip(chunks, current, argc, line); true }
        "kotlin.zip_with_next" => { crate::emitter::hof::emit_zip_with_next(chunks, current, argc, line); true }
        "kotlin.unzip" => { crate::emitter::hof::emit_unzip(chunks, current, argc, line); true }
        "kotlin.with_index" => { crate::emitter::hof::emit_with_index(chunks, current, argc, line); true }
        "kotlin.partition" => { crate::emitter::hof::emit_partition(chunks, current, argc, line); true }
        "kotlin.average" => { crate::emitter::hof::emit_average(chunks, current, argc, line); true }
        "kotlin.take_last" => { crate::emitter::hof::emit_take_last(chunks, current, argc, line); true }
        "kotlin.drop_last" => { crate::emitter::hof::emit_drop_last(chunks, current, argc, line); true }
        "kotlin.flatten" => { crate::emitter::hof::emit_flatten(chunks, current, argc, line); true }
        "kotlin.or_empty" => { crate::emitter::hof::emit_or_empty(chunks, current, argc, line); true }
        "kotlin.sorted_descending" => { crate::emitter::hof::emit_sorted_descending(chunks, current, argc, line); true }
        "kotlin.to_map" => { crate::emitter::hof::emit_to_map(chunks, current, argc, line); true }
        "kotlin.grouping_by" => { crate::emitter::hof::emit_grouping_by(chunks, current, argc, line); true }
        "kotlin.each_count" => { crate::emitter::hof::emit_each_count(chunks, current, argc, line); true }
        "kotlin.grouping_fold" => { crate::emitter::hof::emit_grouping_fold(chunks, current, argc, line); true }
        "kotlin.grouping_reduce" => { crate::emitter::hof::emit_grouping_reduce(chunks, current, argc, line); true }
        "kotlin.grouping_aggregate" => { crate::emitter::hof::emit_grouping_aggregate(chunks, current, argc, line); true }
        "kotlin.list_of_not_null" => { crate::emitter::hof::emit_list_of_not_null(chunks, current, argc, line); true }
        "kotlin.zeroed_array" => { crate::emitter::hof::emit_zeroed_array(chunks, current, argc, line); true }
        "kotlin.map_filter" => { crate::emitter::maps::emit_map_filter(chunks, current, argc, line); true }
        "kotlin.map_filter_not" => { crate::emitter::maps::emit_map_filter_not(chunks, current, argc, line); true }
        "kotlin.filter_keys" => { crate::emitter::maps::emit_filter_keys(chunks, current, argc, line); true }
        "kotlin.filter_values" => { crate::emitter::maps::emit_filter_values(chunks, current, argc, line); true }
        "kotlin.map_values_transform" => { crate::emitter::maps::emit_map_values_transform(chunks, current, argc, line); true }
        "kotlin.map_keys_transform" => { crate::emitter::maps::emit_map_keys_transform(chunks, current, argc, line); true }
        "kotlin.map_to_list" => { crate::emitter::maps::emit_map_to_list(chunks, current, argc, line); true }
        "kotlin.get_value" => { crate::emitter::maps::emit_get_value(chunks, current, argc, line); true }
        "kotlin.with_default" => { crate::emitter::maps::emit_with_default(chunks, current, argc, line); true }
        "kotlin.get_or_put" => { crate::emitter::maps::emit_get_or_put(chunks, current, argc, line); true }
        "kotlin.put_all" => { crate::emitter::maps::emit_put_all(chunks, current, argc, line); true }
        "kotlin.put_if_absent" => { crate::emitter::maps::emit_put_if_absent(chunks, current, argc, line); true }
        "kotlin.copy_map" => { crate::emitter::maps::emit_copy_map(chunks, current, argc, line); true }
        "kotlin.to_sorted_map" => { crate::emitter::maps::emit_to_sorted_map(chunks, current, argc, line); true }
        "kotlin.plus" => { crate::emitter::maps::emit_plus(chunks, current, argc, line); true }
        "kotlin.minus" => { crate::emitter::maps::emit_minus(chunks, current, argc, line); true }
        "kotlin.entries" => { crate::emitter::collections::emit_entries(chunks, current, argc, line); true }
        "kotlin.as_list" => { crate::emitter::collections::emit_dict_as_list(chunks, current, line); true }
        "kotlin.value_eq" => { crate::emitter::equality::emit_value_eq(chunks, current, line); true }
        "kotlin.ref_eq" => { crate::emitter::equality::emit_ref_eq(chunks, current, line); true }
        // Identity at RUNTIME — the marker's whole job is carrying the
        // Double static type out of the erased `toDouble(unit)` call.
        "kotlin.as_double" => true,
        "kotlin.duration_whole" => {
            crate::emitter::time::emit_duration_whole(chunks, current, argc, line);
            true
        }
        "kotlin.duration_str" => {
            crate::emitter::time::emit_duration_str(chunks, current, argc, line);
            true
        }
        "kotlin.print" => {
            crate::emitter::tostring::emit_print(chunks, current, argc, line);
            true
        }
        "kotlin.print_double" => {
            crate::emitter::numbers::emit_print_double(chunks, current, argc, line);
            true
        }
        "kotlin.is_infinite" => {
            crate::emitter::numbers::emit_is_infinite(chunks, current, argc, line);
            true
        }
        "kotlin.to_int_or_null" => {
            crate::emitter::numbers::emit_to_int_or_null(chunks, current, argc, line);
            true
        }
        "kotlin.to_double_or_null" => {
            crate::emitter::numbers::emit_to_double_or_null(chunks, current, argc, line);
            true
        }
        "kotlin.cmp_lt0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Lt,
                line,
            );
            true
        }
        "kotlin.cmp_gt0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Gt,
                line,
            );
            true
        }
        "kotlin.cmp_le0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Le,
                line,
            );
            true
        }
        "kotlin.cmp_ge0" => {
            crate::emitter::numbers::emit_compare_zero(
                chunks,
                current,
                crate::emitter::numbers::CompareZero::Ge,
                line,
            );
            true
        }
        "kotlin.tostring" => {
            crate::emitter::tostring::emit_to_string(chunks, current, line);
            true
        }
        "kotlin.join_to_string" => {
            crate::emitter::tostring::emit_join_to_string(chunks, current, argc, line);
            true
        }
        "kotlin.add" => {
            crate::emitter::collections::emit_add(chunks, current, argc, line);
            true
        }
        "kotlin.set_add" => {
            crate::emitter::collections::emit_set_add(chunks, current, argc, line);
            true
        }
        "kotlin.set_size" => {
            crate::emitter::collections::emit_set_size(chunks, current, argc, line);
            true
        }
        "kotlin.to_set" => {
            crate::emitter::collections::emit_to_set(chunks, current, argc, line);
            true
        }
        "kotlin.to_list" => {
            crate::emitter::collections::emit_to_list(chunks, current, argc, line);
            true
        }
        "kotlin.add_all" => {
            crate::emitter::collections::emit_add_all(chunks, current, argc, line);
            true
        }
        "kotlin.set_union" => {
            crate::emitter::collections::emit_set_union(chunks, current, argc, line);
            true
        }
        "kotlin.set_intersect" => {
            crate::emitter::collections::emit_set_intersect(chunks, current, argc, line);
            true
        }
        "kotlin.set_subtract" => {
            crate::emitter::collections::emit_set_subtract(chunks, current, argc, line);
            true
        }
        "kotlin.remove_all" => {
            crate::emitter::collections::emit_remove_all(chunks, current, argc, line);
            true
        }
        "kotlin.retain_all" => {
            crate::emitter::collections::emit_retain_all(chunks, current, argc, line);
            true
        }
        "kotlin.contains_all" => {
            crate::emitter::collections::emit_contains_all(chunks, current, argc, line);
            true
        }
        "kotlin.is_empty" => {
            crate::emitter::collections::emit_is_empty(chunks, current, argc, line);
            true
        }
        "kotlin.is_not_empty" => {
            crate::emitter::collections::emit_is_not_empty(chunks, current, argc, line);
            true
        }
        "kotlin.if_empty" => {
            crate::emitter::hof::emit_if_empty(chunks, current, argc, line);
            true
        }
        "kotlin.throw_expr" => {
            crate::emitter::nullability::emit_throw_expr(chunks, current, argc, line);
            true
        }
        "kotlin.not_null_assert" => {
            crate::emitter::nullability::emit_not_null_assert(chunks, current, argc, line);
            true
        }
        "kotlin.class_of" => {
            crate::emitter::nullability::emit_class_of(chunks, current, argc, line);
            true
        }
        "kotlin.error" => {
            crate::emitter::nullability::emit_error(chunks, current, argc, line);
            true
        }
        "kotlin.exception" => {
            crate::emitter::nullability::emit_exception(chunks, current, argc, "Exception", line);
            true
        }
        "kotlin.illegal_argument_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "IllegalArgumentException",
                line,
            );
            true
        }
        "kotlin.illegal_state_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "IllegalStateException",
                line,
            );
            true
        }
        "kotlin.null_pointer_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "NullPointerException",
                line,
            );
            true
        }
        "kotlin.index_out_of_bounds_exception" => {
            crate::emitter::nullability::emit_exception(
                chunks,
                current,
                argc,
                "IndexOutOfBoundsException",
                line,
            );
            true
        }
        "kotlin.require" => {
            crate::emitter::nullability::emit_precondition(
                chunks,
                current,
                argc,
                "IllegalArgumentException",
                line,
            );
            true
        }
        "kotlin.check" => {
            crate::emitter::nullability::emit_precondition(
                chunks,
                current,
                argc,
                "IllegalStateException",
                line,
            );
            true
        }
        "kotlin.require_not_null" => {
            crate::emitter::nullability::emit_precondition_not_null(
                chunks,
                current,
                argc,
                "IllegalArgumentException",
                line,
            );
            true
        }
        "kotlin.check_not_null" => {
            crate::emitter::nullability::emit_precondition_not_null(
                chunks,
                current,
                argc,
                "IllegalStateException",
                line,
            );
            true
        }
        _ => false,
    }
}
