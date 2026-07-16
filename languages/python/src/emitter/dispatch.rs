//! Auto-extracted `python.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "python.extend" => {
            crate::emitter::collections_adapter::emit_extend(chunks, current, line)
        }
        "python.get" => {
            crate::emitter::collections_adapter::emit_get(chunks, current, argc, line)
        }
        "python.pop" => {
            crate::emitter::collections_adapter::emit_pop(chunks, current, argc, line)
        }
        "python.index" => {
            crate::emitter::collections_adapter::emit_index(chunks, current, argc, line)
        }
        "python.from_end" => {
            crate::emitter::collections_adapter::emit_from_end(chunks, current, argc, line)
        }
        "python.contains" => {
            crate::emitter::collections_adapter::emit_contains(chunks, current, line)
        }
        "python.next" => {
            crate::emitter::collections_adapter::emit_pynext(chunks, current, argc, line)
        }
        "python.it_reduce" => {
            crate::emitter::itertools_adapter::emit_reduce(chunks, current, argc, line)
        }
        "python.it_filterfalse" => {
            crate::emitter::itertools_adapter::emit_filterfalse(chunks, current, argc, line)
        }
        "python.it_takewhile" => {
            crate::emitter::itertools_adapter::emit_takewhile(chunks, current, argc, line)
        }
        "python.it_dropwhile" => {
            crate::emitter::itertools_adapter::emit_dropwhile(chunks, current, argc, line)
        }
        "python.it_zip_longest" => {
            crate::emitter::itertools_adapter::emit_zip_longest(chunks, current, argc, line)
        }
        "python.op_truth" => {
            crate::emitter::itertools_adapter::emit_op_truth(chunks, current, argc, line)
        }
        "python.op_not" => {
            crate::emitter::itertools_adapter::emit_op_not(chunks, current, argc, line)
        }
        "python.op_eq" => {
            crate::emitter::itertools_adapter::emit_op_eq(chunks, current, argc, line)
        }
        "python.op_ne" => {
            crate::emitter::itertools_adapter::emit_op_ne(chunks, current, argc, line)
        }
        "python.op_pos" => {
            crate::emitter::itertools_adapter::emit_op_pos(chunks, current, argc, line)
        }
        "python.op_abs" => {
            crate::emitter::itertools_adapter::emit_op_abs(chunks, current, argc, line)
        }
        "python.op_inv" => {
            crate::emitter::itertools_adapter::emit_op_inv(chunks, current, argc, line)
        }
        "python.op_and" => {
            crate::emitter::itertools_adapter::emit_op_and(chunks, current, argc, line)
        }
        "python.op_or" => {
            crate::emitter::itertools_adapter::emit_op_or(chunks, current, argc, line)
        }
        "python.op_xor" => {
            crate::emitter::itertools_adapter::emit_op_xor(chunks, current, argc, line)
        }
        "python.op_lshift" => {
            crate::emitter::itertools_adapter::emit_op_lshift(chunks, current, argc, line)
        }
        "python.op_rshift" => {
            crate::emitter::itertools_adapter::emit_op_rshift(chunks, current, argc, line)
        }
        "python.op_getitem" => {
            crate::emitter::itertools_adapter::emit_op_getitem(chunks, current, argc, line)
        }
        "python.op_setitem" => {
            crate::emitter::itertools_adapter::emit_op_setitem(chunks, current, argc, line)
        }
        "python.op_concat" => {
            crate::emitter::itertools_adapter::emit_op_concat(chunks, current, argc, line)
        }
        "python.it_chain" => {
            crate::emitter::itertools_adapter::emit_chain(chunks, current, argc, line)
        }
        "python.it_repeat" => {
            crate::emitter::itertools_adapter::emit_repeat(chunks, current, argc, line)
        }
        "python.it_count" => {
            crate::emitter::itertools_adapter::emit_count(chunks, current, argc, line)
        }
        "python.it_cycle" => {
            crate::emitter::itertools_adapter::emit_cycle(chunks, current, argc, line)
        }
        "python.it_islice" => {
            crate::emitter::itertools_adapter::emit_islice(chunks, current, argc, line)
        }
        "python.it_accumulate" => {
            crate::emitter::itertools_adapter::emit_accumulate(chunks, current, argc, line)
        }
        "python.it_pairwise" => {
            crate::emitter::itertools_adapter::emit_pairwise(chunks, current, argc, line)
        }
        "python.it_batched" => {
            crate::emitter::itertools_adapter::emit_batched(chunks, current, argc, line)
        }
        "python.it_tee" => {
            crate::emitter::itertools_adapter::emit_tee(chunks, current, argc, line)
        }
        "python.time_gmtime" => {
            crate::emitter::time_adapter::emit_gmtime(chunks, current, argc, line)
        }
        "python.time_struct_time" => {
            crate::emitter::time_adapter::emit_struct_time(chunks, current, argc, line)
        }
        "python.time_mktime" => {
            crate::emitter::time_adapter::emit_mktime(chunks, current, argc, line)
        }
        "python.time_clock_seconds" => {
            crate::emitter::time_adapter::emit_clock_seconds(chunks, current, argc, line)
        }
        "python.time_clock_ns" => {
            crate::emitter::time_adapter::emit_clock_ns(chunks, current, argc, line)
        }
        "python.array_new" => {
            crate::emitter::array_adapter::emit_array_new(chunks, current, argc, line)
        }
        "python.array_tolist" => {
            crate::emitter::array_adapter::emit_tolist(chunks, current, argc, line)
        }
        "python.array_buffer_info" => {
            crate::emitter::array_adapter::emit_buffer_info(chunks, current, argc, line)
        }
        "python.array_frombytes" => {
            crate::emitter::array_adapter::emit_frombytes(chunks, current, argc, line)
        }
        "python.heapify" => {
            crate::emitter::heapq_adapter::emit_heapify(chunks, current, argc, line)
        }
        "python.heappush" => {
            crate::emitter::heapq_adapter::emit_heappush(chunks, current, argc, line)
        }
        "python.heappop" => {
            crate::emitter::heapq_adapter::emit_heappop(chunks, current, argc, line)
        }
        "python.heapreplace" => {
            crate::emitter::heapq_adapter::emit_heapreplace(chunks, current, argc, line)
        }
        "python.heappushpop" => {
            crate::emitter::heapq_adapter::emit_heappushpop(chunks, current, argc, line)
        }
        "python.nsmallest" => {
            crate::emitter::heapq_adapter::emit_nsmallest(chunks, current, argc, line)
        }
        "python.nlargest" => {
            crate::emitter::heapq_adapter::emit_nlargest(chunks, current, argc, line)
        }
        "python.heapmerge" => {
            crate::emitter::heapq_adapter::emit_merge(chunks, current, argc, line)
        }
        "python.bisect_left" => {
            crate::emitter::bisect_adapter::emit_bisect_left(chunks, current, argc, line)
        }
        "python.bisect_right" => {
            crate::emitter::bisect_adapter::emit_bisect_right(chunks, current, argc, line)
        }
        "python.insort_left" => {
            crate::emitter::bisect_adapter::emit_insort_left(chunks, current, argc, line)
        }
        "python.insort_right" => {
            crate::emitter::bisect_adapter::emit_insort_right(chunks, current, argc, line)
        }
        "python.stat_quantiles" => {
            crate::emitter::statistics_adapter::emit_quantiles(chunks, current, argc, line)
        }
        "python.stat_median_grouped" => {
            crate::emitter::statistics_adapter::emit_median_grouped(chunks, current, argc, line)
        }
        "python.stat_mode" => {
            crate::emitter::statistics_adapter::emit_mode(chunks, current, argc, line)
        }
        "python.stat_multimode" => {
            crate::emitter::statistics_adapter::emit_multimode(chunks, current, argc, line)
        }
        "python.stat_mean" => {
            crate::emitter::statistics_adapter::emit_mean(chunks, current, argc, line)
        }
        "python.stat_median" => {
            crate::emitter::statistics_adapter::emit_median(chunks, current, argc, line)
        }
        "python.stat_median_low" => {
            crate::emitter::statistics_adapter::emit_median_low(chunks, current, argc, line)
        }
        "python.stat_median_high" => {
            crate::emitter::statistics_adapter::emit_median_high(chunks, current, argc, line)
        }
        "python.stat_variance" => {
            crate::emitter::statistics_adapter::emit_variance(chunks, current, argc, line)
        }
        "python.stat_pvariance" => {
            crate::emitter::statistics_adapter::emit_pvariance(chunks, current, argc, line)
        }
        "python.stat_stdev" => {
            crate::emitter::statistics_adapter::emit_stdev(chunks, current, argc, line)
        }
        "python.stat_pstdev" => {
            crate::emitter::statistics_adapter::emit_pstdev(chunks, current, argc, line)
        }
        "python.stat_harmonic_mean" => {
            crate::emitter::statistics_adapter::emit_harmonic_mean(chunks, current, argc, line)
        }
        "python.stat_geometric_mean" => {
            crate::emitter::statistics_adapter::emit_geometric_mean(chunks, current, argc, line)
        }
        "python.date_new" => {
            crate::emitter::datetime_adapter::emit_date_new(chunks, current, argc, line)
        }
        "python.time_new" => {
            crate::emitter::datetime_adapter::emit_time_new(chunks, current, argc, line)
        }
        "python.datetime_new" => {
            crate::emitter::datetime_adapter::emit_datetime_new(chunks, current, argc, line)
        }
        "python.timedelta_new" => {
            crate::emitter::datetime_adapter::emit_timedelta_new(chunks, current, argc, line)
        }
        "python.timezone_new" => {
            crate::emitter::datetime_adapter::emit_timezone_new(chunks, current, argc, line)
        }
        "python.total_seconds" => {
            crate::emitter::datetime_adapter::emit_total_seconds(chunks, current, argc, line)
        }
        "python.utcoffset" => {
            crate::emitter::datetime_adapter::emit_utcoffset(chunks, current, argc, line)
        }
        "python.timezone_utc" => {
            crate::emitter::datetime_adapter::emit_timezone_utc(chunks, current, argc, line)
        }
        "python.timedelta_resolution" => {
            crate::emitter::datetime_adapter::emit_timedelta_resolution(chunks, current, argc, line)
        }
        "python.date_min" => {
            crate::emitter::datetime_adapter::emit_date_min(chunks, current, argc, line)
        }
        "python.date_max" => {
            crate::emitter::datetime_adapter::emit_date_max(chunks, current, argc, line)
        }
        "python.toordinal" => {
            crate::emitter::datetime_adapter::emit_toordinal(chunks, current, argc, line)
        }
        "python.fromordinal" => {
            crate::emitter::datetime_adapter::emit_fromordinal(chunks, current, argc, line)
        }
        "python.fromtimestamp" => {
            crate::emitter::datetime_adapter::emit_fromtimestamp(chunks, current, argc, line)
        }
        "python.timestamp" => {
            crate::emitter::datetime_adapter::emit_timestamp(chunks, current, argc, line)
        }
        "python.date_fromisoformat" => {
            crate::emitter::datetime_adapter::emit_date_fromisoformat(chunks, current, argc, line)
        }
        "python.datetime_fromisoformat" => {
            crate::emitter::datetime_adapter::emit_datetime_fromisoformat(
                chunks, current, argc, line,
            )
        }
        "python.time_fromisoformat" => {
            crate::emitter::datetime_adapter::emit_time_fromisoformat(chunks, current, argc, line)
        }
        "python.dt_now" => {
            crate::emitter::datetime_adapter::emit_now(chunks, current, argc, line)
        }
        "python.dt_today" => {
            crate::emitter::datetime_adapter::emit_today(chunks, current, argc, line)
        }
        "python.dt_combine" => {
            crate::emitter::datetime_adapter::emit_combine(chunks, current, argc, line)
        }
        "python.dt_date_method" => {
            crate::emitter::datetime_adapter::emit_date_method(chunks, current, argc, line)
        }
        "python.dt_time_method" => {
            crate::emitter::datetime_adapter::emit_time_method(chunks, current, argc, line)
        }
        "python.timetuple" => {
            crate::emitter::datetime_adapter::emit_timetuple(chunks, current, argc, line)
        }
        "python.dt_pad" => {
            crate::emitter::datetime_adapter::emit_dt_pad(chunks, current, argc, line)
        }
        "python.dt_replace" => {
            crate::emitter::datetime_adapter::emit_dt_replace(chunks, current, argc, line)
        }
        "python.cal_weekday" => {
            crate::emitter::datetime_adapter::emit_cal_weekday(chunks, current, argc, line)
        }
        "python.cal_isleap" => {
            crate::emitter::datetime_adapter::emit_cal_isleap(chunks, current, argc, line)
        }
        "python.cal_monthrange" => {
            crate::emitter::datetime_adapter::emit_cal_monthrange(chunks, current, argc, line)
        }
        "python.dt_isoformat" => {
            crate::emitter::datetime_adapter::emit_isoformat(chunks, current, argc, line)
        }
        "python.date_weekday" => {
            crate::emitter::datetime_adapter::emit_date_weekday(chunks, current, argc, line)
        }
        "python.float_repr" => {
            crate::emitter::float_adapter::emit_float_repr(chunks, current, argc, line)
        }
        "python.gen_send" => {
            crate::emitter::collections_adapter::emit_gen_send(chunks, current, argc, line)
        }
        "python.gen_throw" => {
            crate::emitter::collections_adapter::emit_gen_throw(chunks, current, argc, line)
        }
        "python.add" => {
            crate::emitter::collections_adapter::emit_add(chunks, current, line)
        }
        "python.remove" => {
            crate::emitter::collections_adapter::emit_remove(chunks, current, line)
        }
        "python.discard" => {
            crate::emitter::collections_adapter::emit_discard(chunks, current, line)
        }
        "python.copy" => {
            crate::emitter::collections_adapter::emit_copy(chunks, current, line)
        }
        "python.update" => {
            crate::emitter::collections_adapter::emit_update(chunks, current, line)
        }
        "python.intersection_update" => {
            crate::emitter::collections_adapter::emit_intersection_update(
                chunks, current, line,
            )
        }
        "python.difference_update" => {
            crate::emitter::collections_adapter::emit_difference_update(
                chunks, current, line,
            )
        }
        "python.symmetric_difference_update" => {
            crate::emitter::collections_adapter::emit_symmetric_difference_update(
                chunks, current, line,
            )
        }
        "python.clear" => {
            crate::emitter::collections_adapter::emit_clear(chunks, current, line)
        }
        "python.length" => {
            crate::emitter::collections_adapter::emit_length(chunks, current, line)
        }
        "python.str" => {
            crate::emitter::runtime_adapter::emit_str(chunks, current, argc, line)
        }
        "python.repr" => {
            crate::emitter::runtime_adapter::emit_repr(chunks, current, argc, line)
        }
        "python.print" => {
            crate::emitter::runtime_adapter::emit_print(chunks, current, argc, line)
        }
        "python.bytes_decode" => {
            crate::emitter::runtime_adapter::emit_bytes_decode(chunks, current, argc, line)
        }
        "python.pyneg" => {
            crate::emitter::runtime_adapter::emit_pyneg(chunks, current, line)
        }
        "python.pylt" => crate::emitter::runtime_adapter::emit_pylt(chunks, current, line),
        "python.pygt" => crate::emitter::runtime_adapter::emit_pygt(chunks, current, line),
        "python.pyle" => crate::emitter::runtime_adapter::emit_pyle(chunks, current, line),
        "python.pyge" => crate::emitter::runtime_adapter::emit_pyge(chunks, current, line),
        "python.pyadd" => {
            crate::emitter::runtime_adapter::emit_pyadd(chunks, current, line)
        }
        "python.pymul" => {
            crate::emitter::runtime_adapter::emit_pymul(chunks, current, line)
        }
        "python.pysub" => {
            crate::emitter::runtime_adapter::emit_pysub(chunks, current, line)
        }
        "python.pytruediv" => {
            crate::emitter::runtime_adapter::emit_pytruediv(chunks, current, line)
        }
        "python.pyfloordiv" => {
            crate::emitter::runtime_adapter::emit_pyfloordiv(chunks, current, line)
        }
        "python.pymod" => {
            crate::emitter::runtime_adapter::emit_pymod(chunks, current, line)
        }
        "python.pypow" => {
            crate::emitter::runtime_adapter::emit_pypow(chunks, current, line)
        }
        "python.range" => {
            crate::emitter::runtime_adapter::emit_range(chunks, current, argc, line)
        }
        name if crate::emitter::runtime_adapter::emit_helper(
            name, chunks, current, argc, line,
        ) => {}

        // ── COBOL adapters ──
        _ => return false,
    }
    true
}
