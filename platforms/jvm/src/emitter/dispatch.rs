//! `jvm.*` emit dispatch — the platform's own adapter routing.
//!
//! Mirrors `platforms/dotnet`: an op named `common:jvm.java.<name>` reaches this
//! function through `primitives::dispatch` → `platform_emit_dispatch_for`, so
//! the JDK's behaviour is emitted from the PLATFORM and every JVM language
//! gets it by resolving the tree — no per-language emitter arms, no prelude.

use vybe_compiler::primitives::url::UrlField;
use vybe_compiler::primitives::{
    collections,
    instructions::{core_wasm, host},
    ops, sorted_collection, strings,
};
use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    use crate::emitter::arrays_adapter as arrays;
    use crate::emitter::bitset_adapter as bitset;
    use crate::emitter::collection_adapter as collection;
    use crate::emitter::instant_adapter as instant;
    use crate::emitter::enum_set_adapter as enum_set;
    use crate::emitter::io_adapter as io;
    use crate::emitter::map_adapter as map;
    use crate::emitter::math_adapter as math;
    use crate::emitter::optional_adapter as optional;
    use crate::emitter::random_adapter as random;
    use crate::emitter::stream_adapter as stream;
    use crate::emitter::string_adapter;
    use crate::emitter::stringbuilder_adapter as sb;
    use crate::emitter::stringtokenizer_adapter as st;
    use crate::emitter::system_adapter as system;
    use crate::emitter::url_adapter as url;
    use crate::emitter::uuid_adapter as uuid;
    match name {
        // ── java.util.EnumSet ──
        "jvm.java.enum_set_none_of" => enum_set::emit_none_of(chunks, current, line),
        "jvm.java.enum_set_all_of" => enum_set::emit_all_of(chunks, current, line),
        "jvm.java.enum_set_of" => enum_set::emit_of(chunks, current, argc, line),
        "jvm.java.enum_set_copy_of" => enum_set::emit_copy_of(chunks, current, line),
        "jvm.java.enum_set_complement_of" => enum_set::emit_complement_of(chunks, current, line),
        "jvm.java.enum_set_range" => enum_set::emit_range(chunks, current, line),
        "jvm.java.enum_set_add" => enum_set::emit_add(chunks, current, line),
        "jvm.java.enum_set_add_all" => enum_set::emit_add_all(chunks, current, line),
        "jvm.java.enum_set_contains" => enum_set::emit_contains(chunks, current, line),
        "jvm.java.enum_set_contains_all" => enum_set::emit_contains_all(chunks, current, line),
        "jvm.java.enum_set_remove" => enum_set::emit_remove(chunks, current, line),
        "jvm.java.enum_set_equals" => enum_set::emit_equals(chunks, current, line),
        "jvm.java.enum_set_hash_code" => enum_set::emit_hash_code(chunks, current, line),
        "jvm.java.enum_set_iterator" => enum_set::emit_iterator(chunks, current, line),
        "jvm.java.enum_set_get_class" => enum_set::emit_get_class(chunks, current, line),

        // ── java.lang.Enum class metadata ──
        // `Class.getEnumConstants()`, published by each enum's static
        // initializer. `X.class` is a NAME, so this is how a leaf that is
        // handed only that name reaches the constants.
        "jvm.java.enum_declare" => {
            crate::emitter::enum_adapter::emit_declare(chunks, current, line)
        }
        "jvm.java.enum_constants_of" => {
            crate::emitter::enum_adapter::emit_constants_of(chunks, current, line)
        }

        // ── java.lang.Object ──
        // `String.valueOf` / `println(Object)` / `"" + x` all render through
        // one adapter, so a class's ToString slot is reached identically from
        // every one of them.
        "jvm.java.to_string" => crate::emitter::object_adapter::emit_to_string(chunks, current, line),

        // ── construction ──
        "jvm.java.net.url_new" => url::emit_url_new(chunks, current, argc, line),
        "jvm.java.net.uri_new" => url::emit_uri_new(chunks, current, argc, line),
        "jvm.java.random_new" => random::emit_new(chunks, current, argc, line),
        "jvm.java.io_byte_array_output_stream_new" => {
            io::emit_byte_array_output_stream_new(chunks, current, argc, line)
        }
        "jvm.java.io_byte_array_input_stream_new" => {
            io::emit_byte_array_input_stream_new(chunks, current, argc, line)
        }
        "jvm.java.io_sequence_input_stream_new" => {
            io::emit_sequence_input_stream_new(chunks, current, argc, line)
        }
        "jvm.java.io_print_writer_new" => io::emit_print_writer_new(chunks, current, argc, line),
        "jvm.java.io_passthrough_new" => io::emit_passthrough_new(chunks, current, argc, line),
        "jvm.java.io_string_writer_new" => io::emit_string_writer_new(chunks, current, argc, line),
        "jvm.java.io_string_reader_new" => io::emit_string_reader_new(chunks, current, argc, line),

        // ── components: the SHARED reader in primitives/url.rs ──
        "jvm.java.net.url_protocol" => {
            url::emit_component_getter(chunks, current, UrlField::Scheme, line)
        }
        "jvm.java.net.url_host" => {
            url::emit_component_getter(chunks, current, UrlField::Host, line)
        }
        "jvm.java.net.url_path" => {
            url::emit_component_getter(chunks, current, UrlField::Path, line)
        }
        "jvm.java.net.url_authority" => {
            url::emit_component_getter(chunks, current, UrlField::Netloc, line)
        }
        "jvm.java.net.url_query" => {
            url::emit_nullable_getter(chunks, current, UrlField::Query, line)
        }
        "jvm.java.net.url_ref" => {
            url::emit_nullable_getter(chunks, current, UrlField::Fragment, line)
        }
        "jvm.java.net.url_port" => url::emit_port(chunks, current, line),
        "jvm.java.net.url_default_port" => url::emit_default_port(chunks, current, line),
        "jvm.java.net.url_file" => url::emit_file(chunks, current, line),
        "jvm.java.net.url_user_info" => url::emit_user_info(chunks, current, line),

        // ── identity and text ──
        "jvm.java.net.url_to_string" => url::emit_to_string(chunks, current, line),
        "jvm.java.net.url_to_uri" => url::emit_url_to_uri(chunks, current, line),
        "jvm.java.net.url_equals" => url::emit_equals(chunks, current, line),
        "jvm.java.net.url_hash" => url::emit_hash(chunks, current, line),
        "jvm.java.net.url_same_file" => url::emit_same_file(chunks, current, line),
        "jvm.java.net.url_encode" => url::emit_url_encode(chunks, current, line),
        "jvm.java.net.url_decode" => url::emit_url_decode(chunks, current, line),

        // ── java.net.URI's relational surface ──
        "jvm.java.net.uri_ssp" => url::emit_ssp(chunks, current, line),
        "jvm.java.net.uri_is_absolute" => url::emit_is_absolute(chunks, current, line),
        "jvm.java.net.uri_is_opaque" => url::emit_is_opaque(chunks, current, line),
        "jvm.java.net.uri_to_url" => url::emit_uri_to_url(chunks, current, line),
        "jvm.java.net.uri_normalize" => url::emit_normalize(chunks, current, line),
        "jvm.java.net.uri_resolve" => url::emit_resolve(chunks, current, line),
        "jvm.java.net.uri_relativize" => url::emit_relativize(chunks, current, line),
        "jvm.java.net.uri_compare_to" => url::emit_compare_to(chunks, current, line),
        "jvm.java.lang.system_get_property" => {
            system::emit_get_property(chunks, current, argc, line);
        }
        "jvm.java.instant_of_epoch_second" => {
            instant::emit_of_epoch_second(chunks, current, argc, line)
        }
        "jvm.java.instant_of_epoch_milli" => instant::emit_of_epoch_milli(chunks, current, line),
        "jvm.java.instant_parse" => instant::emit_parse(chunks, current, line),
        "jvm.java.instant_now" => instant::emit_instant_now(chunks, current, argc, line),
        "jvm.java.clock_fixed" => instant::emit_clock_fixed(chunks, current, line),
        "jvm.java.local_date_of" => instant::emit_local_date_of(chunks, current, line),
        "jvm.java.local_date_parse" => instant::emit_local_date_parse(chunks, current, line),
        "jvm.java.local_time_of" => instant::emit_local_time_of(chunks, current, argc, line),
        "jvm.java.local_time_parse" => instant::emit_local_time_parse(chunks, current, line),
        "jvm.java.local_datetime_of" => {
            instant::emit_local_datetime_of(chunks, current, argc, line)
        }
        "jvm.java.local_datetime_parse" => {
            instant::emit_local_datetime_parse(chunks, current, line)
        }
        "jvm.java.offset_datetime_of" => {
            instant::emit_offset_datetime_of(chunks, current, argc, line)
        }
        "jvm.java.offset_datetime_of_instant" => {
            instant::emit_offset_datetime_of_instant(chunks, current, line)
        }
        "jvm.java.offset_datetime_parse" => {
            instant::emit_offset_datetime_parse(chunks, current, line)
        }
        "jvm.java.zoned_datetime_of" => {
            instant::emit_zoned_datetime_of(chunks, current, argc, line)
        }
        "jvm.java.zoned_datetime_of_instant" => {
            instant::emit_zoned_datetime_of_instant(chunks, current, line)
        }
        "jvm.java.zoned_datetime_of_strict" => {
            instant::emit_zoned_datetime_of_strict(chunks, current, line)
        }
        "jvm.java.zoned_datetime_parse" => {
            instant::emit_zoned_datetime_parse(chunks, current, line)
        }
        "jvm.java.instant_get_epoch_second" => {
            instant::emit_get_epoch_second(chunks, current, line)
        }
        "jvm.java.instant_get_nano" => {
            chunks[current].emit_string_const("nano", line);
            host::emit(&mut chunks[current], "ecma:object", "get", 2, line);
        }
        "jvm.java.instant_to_epoch_milli" => instant::emit_to_epoch_milli(chunks, current, line),
        "jvm.java.instant_plus_seconds" => instant::emit_plus_seconds(chunks, current, 1.0, line),
        "jvm.java.instant_minus_seconds" => instant::emit_plus_seconds(chunks, current, -1.0, line),
        "jvm.java.instant_plus_millis" => instant::emit_plus_millis(chunks, current, 1.0, line),
        "jvm.java.instant_minus_millis" => instant::emit_plus_millis(chunks, current, -1.0, line),
        "jvm.java.instant_plus_nanos" => instant::emit_plus_nanos(chunks, current, 1.0, line),
        "jvm.java.instant_minus_nanos" => instant::emit_plus_nanos(chunks, current, -1.0, line),
        "jvm.java.instant_compare_to" => instant::emit_compare(chunks, current, line),
        "jvm.java.instant_is_before" => instant::emit_is_before_after(chunks, current, false, line),
        "jvm.java.instant_is_after" => instant::emit_is_before_after(chunks, current, true, line),
        "jvm.java.instant_equals" => instant::emit_equals(chunks, current, line),
        "jvm.java.instant_to_string" => instant::emit_to_string(chunks, current, line),
        "jvm.java.duration_of_hours" => instant::emit_duration_hours(chunks, current, line),
        "jvm.java.duration_of_minutes" => instant::emit_duration_minutes(chunks, current, line),
        "jvm.java.duration_of_seconds" => instant::emit_duration_seconds(chunks, current, line),
        "jvm.java.duration_parse" => instant::emit_duration_parse(chunks, current, line),
        "jvm.java.duration_between" => instant::emit_duration_between(chunks, current, line),
        "jvm.java.chrono_days_between" => {
            instant::emit_chrono_between(chunks, current, 86400.0, line)
        }
        "jvm.java.chrono_weeks_between" => {
            instant::emit_chrono_between(chunks, current, 604800.0, line)
        }
        "jvm.java.chrono_months_between" => {
            instant::emit_chrono_between(chunks, current, 2592000.0, line)
        }
        "jvm.java.period_of_days" => instant::emit_period_of_days(chunks, current, line),
        "jvm.java.period_of_months" => instant::emit_period_of_months(chunks, current, line),
        "jvm.java.period_get_days" => instant::emit_period_get_days(chunks, current, line),
        "jvm.java.period_get_months" => instant::emit_period_get_months(chunks, current, line),
        "jvm.java.period_get_years" => instant::emit_period_get_years(chunks, current, line),
        "jvm.java.period_between" => instant::emit_period_between(chunks, current, line),
        "jvm.java.zone_offset_of_hours" => instant::emit_zone_offset_hours(chunks, current, line),
        "jvm.java.zone_id_of" => instant::emit_zone_id_utc(chunks, current, line),
        "jvm.java.zone_id_system_default" => {
            instant::emit_zone_id_system_default(chunks, current, line)
        }
        "jvm.java.zone_id_short_ids" => instant::emit_zone_id_short_ids(chunks, current, line),
        "jvm.java.zone_id_from" => instant::emit_zone_id_from(chunks, current, line),
        "jvm.java.zone_id_of_offset" => instant::emit_zone_id_of_offset(chunks, current, line),
        "jvm.java.zone_normalized" => instant::emit_zone_normalized(chunks, current, line),
        "jvm.java.zone_display_name" => {
            instant::emit_zone_display_name(chunks, current, argc, line)
        }
        "jvm.java.zone_rules_fixed" => instant::emit_zone_rules_fixed(chunks, current, line),
        "jvm.java.zone_rules_get_offset" => {
            instant::emit_zone_rules_get_offset(chunks, current, line)
        }
        "jvm.java.zone_offset_total_seconds" => {
            instant::emit_get_total_seconds(chunks, current, line)
        }
        "jvm.java.zone_compare_to" => instant::emit_zone_compare_to(chunks, current, line),
        "jvm.java.zone_hash_code" => instant::emit_zone_hash_code(chunks, current, line),
        "jvm.java.instant_with_offset" => instant::emit_with_offset(chunks, current, line),
        "jvm.java.instant_with_zone" => instant::emit_with_zone_same_instant(chunks, current, line),
        "jvm.java.instant_get_offset" => instant::emit_get_offset(chunks, current, line),
        "jvm.java.instant_get_zone" => instant::emit_get_zone(chunks, current, line),
        "jvm.java.instant_get_year" => {
            instant::emit_component(chunks, current, "getUTCFullYear", false, line)
        }
        "jvm.java.instant_get_month" => {
            instant::emit_component(chunks, current, "getUTCMonth", true, line)
        }
        "jvm.java.instant_get_day" => {
            instant::emit_component(chunks, current, "getUTCDate", false, line)
        }
        "jvm.java.instant_get_hour" => {
            instant::emit_component(chunks, current, "getUTCHours", false, line)
        }
        "jvm.java.instant_get_minute" => {
            instant::emit_component(chunks, current, "getUTCMinutes", false, line)
        }
        "jvm.java.instant_get_second" => {
            instant::emit_component(chunks, current, "getUTCSeconds", false, line)
        }
        "jvm.java.instant_to_local_date" => instant::emit_local_date_string(chunks, current, line),
        "jvm.java.time_to_string" => instant::emit_time_to_string(chunks, current, line),
        "jvm.java.time_format" => instant::emit_time_format(chunks, current, line),
        "jvm.java.time_plus_days" => {
            instant::emit_time_plus_unit(chunks, current, 1.0, 86400.0, line)
        }
        "jvm.java.time_minus_days" => {
            instant::emit_time_plus_unit(chunks, current, -1.0, 86400.0, line)
        }
        "jvm.java.time_plus_weeks" => {
            instant::emit_time_plus_unit(chunks, current, 1.0, 604800.0, line)
        }
        "jvm.java.time_plus_months" => instant::emit_time_plus_months(chunks, current, 1.0, line),
        "jvm.java.time_minus_months" => instant::emit_time_plus_months(chunks, current, -1.0, line),
        "jvm.java.time_plus_hours" => {
            instant::emit_time_plus_unit(chunks, current, 1.0, 3600.0, line)
        }
        "jvm.java.time_minus_hours" => {
            instant::emit_time_plus_unit(chunks, current, -1.0, 3600.0, line)
        }
        "jvm.java.time_plus_minutes" => {
            instant::emit_time_plus_unit(chunks, current, 1.0, 60.0, line)
        }
        "jvm.java.time_minus_minutes" => {
            instant::emit_time_plus_unit(chunks, current, -1.0, 60.0, line)
        }
        "jvm.java.time_plus_seconds" => {
            instant::emit_time_plus_unit(chunks, current, 1.0, 1.0, line)
        }
        "jvm.java.time_minus_seconds" => {
            instant::emit_time_plus_unit(chunks, current, -1.0, 1.0, line)
        }
        "jvm.java.time_with_year" => {
            instant::emit_time_with_year_or_month(chunks, current, true, line)
        }
        "jvm.java.time_with_month" => {
            instant::emit_time_with_year_or_month(chunks, current, false, line)
        }
        "jvm.java.time_with_day" => {
            instant::emit_time_with_field(chunks, current, "setUTCDate", false, line)
        }
        "jvm.java.time_with_hour" => {
            instant::emit_time_with_field(chunks, current, "setUTCHours", false, line)
        }
        "jvm.java.time_with_minute" => {
            instant::emit_time_with_field(chunks, current, "setUTCMinutes", false, line)
        }
        "jvm.java.time_with_second" => {
            instant::emit_time_with_field(chunks, current, "setUTCSeconds", false, line)
        }
        "jvm.java.time_length_of_month" => {
            instant::emit_time_length_of_month(chunks, current, line)
        }
        "jvm.java.time_range_day" => instant::emit_time_range_day(chunks, current, line),
        "jvm.java.time_is_leap_year" => instant::emit_time_is_leap_year(chunks, current, line),
        "jvm.java.time_day_of_year" => instant::emit_time_day_of_year(chunks, current, line),
        "jvm.java.time_day_of_week" => instant::emit_time_day_of_week(chunks, current, line),
        "jvm.java.duration_to_hours" => instant::emit_duration_to_hours(chunks, current, line),
        "jvm.java.duration_to_minutes" => instant::emit_duration_to_minutes(chunks, current, line),
        "jvm.java.duration_to_millis" => instant::emit_duration_to_millis(chunks, current, line),
        "jvm.java.duration_plus_hours" => {
            instant::emit_duration_plus_hours(chunks, current, 1.0, line)
        }
        "jvm.java.duration_minus_hours" => {
            instant::emit_duration_plus_hours(chunks, current, -1.0, line)
        }
        "jvm.java.duration_plus_minutes" => {
            instant::emit_duration_plus_minutes(chunks, current, 1.0, line)
        }
        "jvm.java.duration_minus_minutes" => {
            instant::emit_duration_plus_minutes(chunks, current, -1.0, line)
        }
        "jvm.java.duration_multiplied_by" => {
            instant::emit_duration_multiplied_by(chunks, current, line)
        }
        "jvm.java.duration_negated" => instant::emit_duration_negated(chunks, current, line),
        "jvm.java.duration_is_zero" => instant::emit_duration_is_zero(chunks, current, line),
        "jvm.java.time_with_offset_same_local" => {
            instant::emit_time_with_offset_same_local(chunks, current, line)
        }
        "jvm.java.time_with_zone_same_local" => {
            instant::emit_with_zone_same_local(chunks, current, line)
        }
        "jvm.java.zoned_later_overlap" => instant::emit_overlap_offset(chunks, current, 1, line),
        "jvm.java.zoned_earlier_overlap" => instant::emit_overlap_offset(chunks, current, 2, line),
        "jvm.java.instant_truncated" => instant::emit_truncated(chunks, current, line),
        "jvm.java.instant_hash_code" => instant::emit_hash_code(chunks, current, line),
        "jvm.java.random_set_seed" => random::emit_set_seed(chunks, current, line),
        "jvm.java.random_next_int" => random::emit_next_int(chunks, current, argc, line),
        "jvm.java.random_next_long" => random::emit_next_long(chunks, current, line),
        "jvm.java.random_next_boolean" => random::emit_next_bool(chunks, current, line),
        "jvm.java.random_next_float" => random::emit_next_float(chunks, current, line),
        "jvm.java.random_next_double" => random::emit_next_double(chunks, current, line),
        "jvm.java.random_next_bytes" => random::emit_next_bytes(chunks, current, line),
        "jvm.java.random_split" => random::emit_split(chunks, current, line),
        "jvm.java.random_ints" => random::emit_ints(chunks, current, argc, line),
        "jvm.java.random_longs" => random::emit_longs(chunks, current, argc, line),
        "jvm.java.random_doubles" => random::emit_doubles(chunks, current, argc, line),
        "jvm.java.io_size" => io::emit_size(chunks, current, line),
        "jvm.java.io_output_to_string" => io::emit_output_to_string(chunks, current, argc, line),
        "jvm.java.io_output_write" => io::emit_output_write(chunks, current, argc, line),
        "jvm.java.io_reset_buffer" => io::emit_reset_buffer(chunks, current, line),
        "jvm.java.io_to_byte_array" => io::emit_to_byte_array(chunks, current, line),
        "jvm.java.io_read" => io::emit_read(chunks, current, argc, line),
        "jvm.java.io_available" => io::emit_available(chunks, current, line),
        "jvm.java.io_mark" => io::emit_mark(chunks, current, argc, line),
        "jvm.java.io_reset_pos" => io::emit_reset_pos(chunks, current, line),
        "jvm.java.io_mark_supported" => io::emit_mark_supported(chunks, current, line),
        "jvm.java.io_skip" => io::emit_skip(chunks, current, line),
        "jvm.java.io_writer_print" => {
            io::emit_writer_print(chunks, current, argc, false, false, line)
        }
        "jvm.java.io_writer_println" => {
            io::emit_writer_print(chunks, current, argc, true, false, line)
        }
        "jvm.java.io_writer_write" => {
            io::emit_writer_write(chunks, current, argc, false, false, line)
        }
        "jvm.java.io_writer_append" => {
            io::emit_writer_write(chunks, current, argc, false, true, line)
        }
        "jvm.java.io_writer_newline" => {
            io::emit_writer_write(chunks, current, 1, true, false, line)
        }
        "jvm.java.io_writer_to_string" => io::emit_writer_to_string(chunks, current, line),
        "jvm.java.io_writer_to_char_array" => io::emit_writer_to_char_array(chunks, current, line),
        "jvm.java.io_flush_close" => io::emit_flush_close(chunks, current, line),
        "jvm.java.io_ready" => io::emit_ready(chunks, current, line),
        "jvm.java.io_unread" => io::emit_unread(chunks, current, line),
        "jvm.java.io_read_line" => io::emit_read_line(chunks, current, line),
        "jvm.java.io_get_line_number" => io::emit_get_line_number(chunks, current, line),
        "jvm.java.io_read_utf" => io::emit_read_utf(chunks, current, line),
        "jvm.java.io_false" => {
            chunks[current].emit_bool_const(false, line);
        }
        "jvm.java.string_to_byte_array" => {
            io::emit_string_to_byte_array(chunks, current, argc, line)
        }
        "jvm.java.string_to_char_array" => io::emit_string_to_char_array(chunks, current, line),
        "jvm.java.chars_to_string" => io::emit_chars_to_string(chunks, current, line),
        "jvm.java.int_to_char" => io::emit_int_to_char(chunks, current, line),
        "jvm.java.char_code" => io::emit_char_code(chunks, current, line),
        "jvm.java.byte_array_new" => io::emit_new_filled_array(chunks, current, argc, false, line),
        "jvm.java.char_array_new" => io::emit_new_filled_array(chunks, current, argc, true, line),
        "jvm.java.uuid_from_string" => uuid::emit_from_string(chunks, current, line),
        "jvm.java.uuid_name_from_bytes" => uuid::emit_name_from_bytes(chunks, current, line),
        "jvm.java.uuid_version" => uuid::emit_version(chunks, current, line),
        "jvm.java.uuid_variant" => uuid::emit_variant(chunks, current, line),
        "jvm.java.uuid_most_bits" => uuid::emit_most_bits(chunks, current, line),
        "jvm.java.uuid_least_bits" => uuid::emit_least_bits(chunks, current, line),
        "jvm.java.uuid_compare_to" => uuid::emit_compare_to(chunks, current, line),
        "jvm.java.uuid_hash_code" => uuid::emit_hash_code(chunks, current, line),
        "jvm.java.uuid_new" => uuid::emit_new(chunks, current, argc, line),
        "jvm.java.is_infinite" => {
            host::emit(&mut chunks[current], "ecma:number", "isFinite", 1, line);
            ops::emit_dyn_not(&mut chunks[current], line);
        }
        "jvm.java.signum" => host::emit(&mut chunks[current], "ecma:math", "sign", 1, line),
        "jvm.java.math_scalb" => math::emit_scalb(chunks, current, line),
        "jvm.java.math_ulp" => math::emit_ulp(chunks, current, line),
        "jvm.java.math_get_exponent" => math::emit_get_exponent(chunks, current, line),
        "jvm.java.math_copy_sign" => math::emit_copy_sign(chunks, current, line),
        "jvm.java.math_next_after" => math::emit_next_after(chunks, current, line),
        "jvm.java.math_next_up" => math::emit_next_up(chunks, current, line),
        "jvm.java.math_next_down" => math::emit_next_down(chunks, current, line),
        "jvm.java.math_fma" => math::emit_fma(chunks, current, line),
        "jvm.java.math_expm1" => math::emit_expm1(chunks, current, line),
        "jvm.java.math_log1p" => math::emit_log1p(chunks, current, line),
        "jvm.java.math_to_degrees" => math::emit_to_degrees(chunks, current, line),
        "jvm.java.math_to_radians" => math::emit_to_radians(chunks, current, line),
        "jvm.java.math_ieee_remainder" => math::emit_ieee_remainder(chunks, current, line),
        "jvm.java.math_add_exact" => math::emit_add_exact(chunks, current, line),
        "jvm.java.math_subtract_exact" => math::emit_subtract_exact(chunks, current, line),
        "jvm.java.math_multiply_exact" => math::emit_multiply_exact(chunks, current, line),
        "jvm.java.math_increment_exact" => math::emit_increment_exact(chunks, current, line),
        "jvm.java.math_decrement_exact" => math::emit_decrement_exact(chunks, current, line),
        "jvm.java.math_negate_exact" => math::emit_negate_exact(chunks, current, line),
        "jvm.java.floor_div" => math::emit_floor_div(chunks, current, line),
        "jvm.java.floor_mod" => math::emit_floor_mod(chunks, current, line),
        "jvm.java.identity" => {}
        "jvm.java.char_is_digit" => {
            host::emit(&mut chunks[current], "ecma:char", "isDigit", 1, line);
        }
        "jvm.java.char_is_letter" => {
            host::emit(&mut chunks[current], "ecma:char", "isLetter", 1, line);
        }
        "jvm.java.char_is_alnum" => {
            host::emit(&mut chunks[current], "ecma:char", "isAlnum", 1, line);
        }
        "jvm.java.char_is_upper" => {
            host::emit(&mut chunks[current], "ecma:char", "isUpper", 1, line);
        }
        "jvm.java.char_is_lower" => {
            host::emit(&mut chunks[current], "ecma:char", "isLower", 1, line);
        }
        "jvm.java.char_is_space" => {
            host::emit(&mut chunks[current], "ecma:char", "isSpace", 1, line);
        }
        "jvm.java.char_to_upper" => strings::emit_to_upper(&mut chunks[current], line),
        "jvm.java.char_to_lower" => strings::emit_to_lower(&mut chunks[current], line),
        "jvm.java.char_numeric" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", 1, line);
        }
        "jvm.java.to_binary_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toBinary", 1, line);
        }
        "jvm.java.to_hex_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toHex", 1, line);
        }
        "jvm.java.to_octal_string" => {
            host::emit(&mut chunks[current], "ecma:number", "toOctal", 1, line);
        }
        "jvm.java.parse_int" => {
            host::emit(&mut chunks[current], "ecma:number", "parseInt", argc, line);
        }
        "jvm.java.int_bit_count" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_POPCNT, line);
        }
        "jvm.java.int_leading_zeros" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_CLZ, line);
        }
        "jvm.java.int_trailing_zeros" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_CTZ, line);
        }
        "jvm.java.int_rotate_left" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_ROTL, line);
        }
        "jvm.java.int_rotate_right" => {
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_ROTR, line);
        }
        "jvm.java.int_lowest_one_bit" => {
            let s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, s, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 0);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SUB, line);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_AND, line);
        }
        "jvm.java.int_highest_one_bit" => {
            let s = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, s, line);
            for shift in [1, 2, 4, 8, 16] {
                chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
                chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
                core_wasm::i32_const(&mut chunks[current], line, shift);
                chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SHR_U, line);
                chunks[current].emit_op(vybe_runtime::opcode::Op::I32_OR, line);
                chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, s, line);
            }
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, s, line);
            core_wasm::i32_const(&mut chunks[current], line, 1);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SHR_U, line);
            chunks[current].emit_op(vybe_runtime::opcode::Op::I32_SUB, line);
        }
        "jvm.java.compare" => {
            let b_slot = chunks[current].alloc_scratch(1);
            let a_slot = chunks[current].alloc_scratch(1);
            let result_slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, b_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, a_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, b_slot, line);
            ops::emit_dyn_lt(&mut chunks[current], line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(-1, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, a_slot, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, b_slot, line);
            ops::emit_dyn_gt(&mut chunks[current], line);
            ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_else(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, result_slot, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, result_slot, line);
        }
        "jvm.java.arrays_sort" => arrays::emit_sort(chunks, current, argc, line),
        "jvm.java.arrays_fill" => arrays::emit_fill(chunks, current, argc, line),
        "jvm.java.arrays_copy_of" => arrays::emit_copy_of(chunks, current, line),
        "jvm.java.arrays_copy_of_range" => arrays::emit_copy_of_range(chunks, current, line),
        "jvm.java.arrays_to_string" => arrays::emit_to_string(chunks, current, line),
        "jvm.java.arrays_deep_to_string" => arrays::emit_deep_to_string(chunks, current, line),
        "jvm.java.arrays_equals" => arrays::emit_equals(chunks, current, line),
        "jvm.java.arrays_deep_equals" => arrays::emit_deep_equals(chunks, current, line),
        "jvm.java.arrays_compare" => arrays::emit_compare(chunks, current, line),
        "jvm.java.arrays_compare_unsigned" => arrays::emit_compare_unsigned(chunks, current, line),
        "jvm.java.arrays_mismatch" => arrays::emit_mismatch(chunks, current, line),
        "jvm.java.arrays_set_all" => arrays::emit_set_all(chunks, current, line),
        "jvm.java.arrays_parallel_prefix" => {
            arrays::emit_parallel_prefix(chunks, current, argc, line)
        }
        "jvm.java.arrays_binary_search" => arrays::emit_binary_search(chunks, current, argc, line),
        "jvm.java.arrays_as_list" => arrays::emit_arrays_as_list(chunks, current, argc, line),
        "jvm.java.arrays_hash_code" => arrays::emit_hash_code(chunks, current, line),
        "jvm.java.arrays_deep_hash_code" => arrays::emit_deep_hash_code(chunks, current, line),
        "jvm.java.collections_sort" => arrays::emit_sort(chunks, current, argc, line),
        "jvm.java.collections_reverse" => collections::emit_reverse(chunks, current, line),
        "jvm.java.collections_shuffle" => {
            if argc == 2 {
                chunks[current].emit_op(vybe_runtime::opcode::Op::DROP, line);
            }
            chunks[current].emit_op(vybe_runtime::opcode::Op::DROP, line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        "jvm.java.collections_fill" => arrays::emit_fill(chunks, current, 2, line),
        "jvm.java.collections_copy" => collection::emit_collection_copy(chunks, current, line),
        "jvm.java.collections_add_all" => collection::emit_add_all(chunks, current, argc, line),
        "jvm.java.collections_rotate" => collection::emit_rotate(chunks, current, line),
        "jvm.java.collections_replace_all" => {
            collection::emit_replace_all_values(chunks, current, line)
        }
        "jvm.java.collections_swap" => collection::emit_swap(chunks, current, line),
        "jvm.java.collections_index_of_sublist" => {
            collection::emit_index_of_sublist(chunks, current, false, line)
        }
        "jvm.java.collections_last_index_of_sublist" => {
            collection::emit_index_of_sublist(chunks, current, true, line)
        }
        "jvm.java.collections_min" => {
            collection::emit_collection_extreme(chunks, current, argc, true, line)
        }
        "jvm.java.collections_max" => {
            collection::emit_collection_extreme(chunks, current, argc, false, line)
        }
        "jvm.java.collections_frequency" => {
            collection::emit_collection_frequency(chunks, current, line)
        }
        "jvm.java.collections_disjoint" => {
            collection::emit_collection_disjoint(chunks, current, line)
        }
        "jvm.java.collections_reverse_order" => {
            collection::emit_reverse_order(chunks, current, line)
        }
        "jvm.java.new_set_from_map" => collection::emit_new_set_from_map(chunks, current, line),
        "jvm.java.unmodifiable_list" => collection::emit_mark_immutable_list(chunks, current, line),
        "jvm.java.unmodifiable_set" => collection::emit_mark_immutable_list(chunks, current, line),
        "jvm.java.unmodifiable_map" => map::emit_mark_immutable_map(chunks, current, line),
        "jvm.java.n_copies" => collection::emit_n_copies(chunks, current, line),
        "jvm.java.string_join" => string_adapter::emit_join(chunks, current, argc, line),
        "jvm.java.require_non_null" => {
            string_adapter::emit_require_non_null(chunks, current, argc, line)
        }
        "jvm.java.optional_empty" => optional::emit_empty(chunks, current, line),
        "jvm.java.optional_of" => optional::emit_of(chunks, current, line),
        "jvm.java.optional_of_nullable" => optional::emit_of_nullable(chunks, current, line),
        "jvm.java.stream_empty" => stream::emit_empty(chunks, current, line),
        "jvm.java.stream_of" => stream::emit_of(chunks, current, argc, line),
        "jvm.java.stream_builder" => stream::emit_builder(chunks, current, line),
        "jvm.java.stream_builder_add" => stream::emit_builder_add(chunks, current, line),
        "jvm.java.stream_range" => stream::emit_range(chunks, current, false, line),
        "jvm.java.stream_range_closed" => stream::emit_range(chunks, current, true, line),
        "jvm.java.stream_concat" => stream::emit_concat(chunks, current, line),
        "jvm.java.stream_collect" => stream::emit_collect(chunks, current, line),
        "jvm.java.stream_generate" => stream::emit_generate(chunks, current, line),
        "jvm.java.stream_iterate" => stream::emit_iterate(chunks, current, argc, line),
        "jvm.java.stream_iterate_strict" => {
            stream::emit_iterate_strict(chunks, current, argc, line)
        }
        "jvm.java.stream_count" => stream::emit_count(chunks, current, line),
        "jvm.java.stream_to_array" => stream::emit_to_array(chunks, current, argc, line),
        "jvm.java.stream_sum" => stream::emit_sum(chunks, current, line),
        "jvm.java.stream_map" => stream::emit_map(chunks, current, line),
        "jvm.java.stream_filter" => stream::emit_filter(chunks, current, line),
        "jvm.java.stream_peek" => stream::emit_peek(chunks, current, line),
        "jvm.java.stream_distinct" => stream::emit_distinct(chunks, current, line),
        "jvm.java.stream_flat_map" => stream::emit_flat_map(chunks, current, line),
        "jvm.java.stream_sorted" => stream::emit_sorted(chunks, current, argc, line),
        "jvm.java.stream_limit" => stream::emit_limit(chunks, current, line),
        "jvm.java.stream_skip" => stream::emit_skip(chunks, current, line),
        "jvm.java.stream_take_while" => stream::emit_take_while(chunks, current, line),
        "jvm.java.stream_drop_while" => stream::emit_drop_while(chunks, current, line),
        "jvm.java.stream_find_first" => stream::emit_find_first(chunks, current, line),
        "jvm.java.stream_min" => stream::emit_extreme_value(chunks, current, argc, true, line),
        "jvm.java.stream_max" => stream::emit_extreme_value(chunks, current, argc, false, line),
        "jvm.java.stream_max_value" => stream::emit_max_value(chunks, current, line),
        "jvm.java.stream_average" => stream::emit_average(chunks, current, line),
        "jvm.java.stream_average_value" => stream::emit_average_value(chunks, current, line),
        "jvm.java.stream_any_match" => stream::emit_any_match(chunks, current, line),
        "jvm.java.stream_all_match" => stream::emit_all_match(chunks, current, line),
        "jvm.java.stream_none_match" => stream::emit_none_match(chunks, current, line),
        "jvm.java.stream_reduce" => stream::emit_reduce(chunks, current, argc, line),
        "jvm.java.stream_for_each" => stream::emit_for_each(chunks, current, line),
        "jvm.java.stream_optional_get" => stream::emit_get_optional_value(chunks, current, line),
        "jvm.java.collectors_joining" => {
            stream::emit_collectors_joining(chunks, current, argc, line)
        }
        "jvm.java.collectors_to_list" => stream::emit_collectors_to_list(chunks, current, line),
        "jvm.java.collectors_to_set" => {
            stream::emit_collector_tag(chunks, current, "toSet", 0, line)
        }
        "jvm.java.collectors_to_collection" => {
            stream::emit_collector_tag(chunks, current, "toCollection", 1, line)
        }
        "jvm.java.collectors_counting" => {
            stream::emit_collector_tag(chunks, current, "counting", 0, line)
        }
        "jvm.java.collectors_summing_int" => {
            stream::emit_collector_tag(chunks, current, "summingInt", 1, line)
        }
        "jvm.java.collectors_averaging_int" => {
            stream::emit_collector_tag(chunks, current, "averagingInt", 1, line)
        }
        "jvm.java.collectors_to_map" => {
            stream::emit_collector_tag(chunks, current, "toMap", 2, line)
        }
        "jvm.java.collectors_mapping" => {
            stream::emit_collector_tag(chunks, current, "mapping", 2, line)
        }
        "jvm.java.collectors_filtering" => {
            stream::emit_collector_tag(chunks, current, "filtering", 2, line)
        }
        "jvm.java.collectors_collecting_and_then" => {
            stream::emit_collector_tag(chunks, current, "collectingAndThen", 2, line)
        }
        "jvm.java.collectors_reducing" => {
            stream::emit_collector_tag(chunks, current, "reducing", argc, line)
        }
        "jvm.java.collectors_min_by" => {
            stream::emit_collector_tag(chunks, current, "minBy", 1, line)
        }
        "jvm.java.collectors_max_by" => {
            stream::emit_collector_tag(chunks, current, "maxBy", 1, line)
        }
        "jvm.java.collectors_grouping_by" => stream::emit_collector_tag_with_default_downstream(
            chunks,
            current,
            "groupingBy",
            argc,
            line,
        ),
        "jvm.java.collectors_partitioning_by" => {
            stream::emit_collector_tag_with_default_downstream(
                chunks,
                current,
                "partitioningBy",
                argc,
                line,
            )
        }
        "jvm.java.empty_list" => collections::emit_array_new(chunks, current, 0, line),
        "jvm.java.empty_set" => {
            collections::emit_array_new(chunks, current, 0, line);
            sorted_collection::emit_mark_set_collection(chunks, current, line);
        }
        "jvm.java.bitset_new" => bitset::emit_new(chunks, current, argc, line),
        "jvm.java.bitset_value_of" => bitset::emit_value_of(chunks, current, line),
        "jvm.java.bitset_set" => bitset::emit_set(chunks, current, argc, line),
        "jvm.java.bitset_get" => bitset::emit_get(chunks, current, argc, line),
        "jvm.java.bitset_clear" => bitset::emit_clear(chunks, current, argc, line),
        "jvm.java.bitset_flip" => bitset::emit_flip(chunks, current, argc, line),
        "jvm.java.bitset_cardinality" => bitset::emit_cardinality(chunks, current, line),
        "jvm.java.bitset_length" => bitset::emit_length(chunks, current, line),
        "jvm.java.bitset_size" => bitset::emit_size(chunks, current, line),
        "jvm.java.bitset_is_empty" => bitset::emit_is_empty(chunks, current, line),
        "jvm.java.bitset_next_set_bit" => bitset::emit_next_set_bit(chunks, current, line),
        "jvm.java.bitset_next_clear_bit" => bitset::emit_next_clear_bit(chunks, current, line),
        "jvm.java.bitset_previous_set_bit" => bitset::emit_previous_set_bit(chunks, current, line),
        "jvm.java.bitset_previous_clear_bit" => {
            bitset::emit_previous_clear_bit(chunks, current, line)
        }
        "jvm.java.bitset_and" => bitset::emit_and(chunks, current, line),
        "jvm.java.bitset_or" => bitset::emit_or(chunks, current, line),
        "jvm.java.bitset_xor" => bitset::emit_xor(chunks, current, line),
        "jvm.java.bitset_and_not" => bitset::emit_and_not(chunks, current, line),
        "jvm.java.bitset_intersects" => bitset::emit_intersects(chunks, current, line),
        "jvm.java.bitset_equals" => bitset::emit_equals(chunks, current, line),
        "jvm.java.bitset_clone" => bitset::emit_clone(chunks, current, line),
        "jvm.java.bitset_stream" => bitset::emit_stream(chunks, current, line),
        "jvm.java.bitset_to_array" => bitset::emit_to_array(chunks, current, line),
        "jvm.java.bitset_to_string" => bitset::emit_to_string(chunks, current, line),
        "jvm.java.bitset_hash_code" => bitset::emit_hash_code(chunks, current, line),
        "jvm.java.mutable_list_new" => {
            collection::emit_mutable_list_new(chunks, current, argc, line)
        }
        "jvm.java.copy_on_write_list_new" => {
            collection::emit_copy_on_write_list_new(chunks, current, argc, line)
        }
        "jvm.java.linked_blocking_queue_new" => {
            collection::emit_linked_blocking_queue_new(chunks, current, argc, line)
        }
        "jvm.java.vector_new" => collection::emit_vector_new(chunks, current, argc, line),
        "jvm.java.hash_set_new" => collection::emit_hash_set_new(chunks, current, argc, line),
        "jvm.java.list_of" => collection::emit_list_of(chunks, current, argc, line),
        "jvm.java.list_copy_of" => collection::emit_list_copy_of(chunks, current, line),
        "jvm.java.set_of" => collection::emit_set_of(chunks, current, argc, line),
        "jvm.java.set_copy_of" => collection::emit_set_copy_of(chunks, current, line),
        "jvm.java.sorted_set_new" => collection::emit_sorted_set_new(chunks, current, argc, line),
        "jvm.java.sorted_map_new" => collection::emit_sorted_map_new(chunks, current, argc, line),
        "jvm.java.priority_queue_new" => {
            collection::emit_priority_queue_new(chunks, current, argc, line)
        }
        "jvm.java.collection_passthrough_new" => {
            collection::emit_passthrough_new(chunks, current, argc, line)
        }
        "jvm.java.list_iterator" => collection::emit_list_iterator(chunks, current, argc, line),
        "jvm.java.iterator_next" => collection::emit_iterator_next(chunks, current, line),
        "jvm.java.iterator_has_next" => collection::emit_iterator_has_next(chunks, current, line),
        "jvm.java.iterator_previous" => collection::emit_iterator_previous(chunks, current, line),
        "jvm.java.iterator_has_previous" => {
            collection::emit_iterator_has_previous(chunks, current, line)
        }
        "jvm.java.iterator_next_index" => {
            collection::emit_iterator_index(chunks, current, false, line)
        }
        "jvm.java.iterator_previous_index" => {
            collection::emit_iterator_index(chunks, current, true, line)
        }
        "jvm.java.add" => collection::emit_add(chunks, current, argc, line),
        "jvm.java.get" => collection::emit_get(chunks, current, line),
        "jvm.java.list_set" => collection::emit_set(chunks, current, line),
        "jvm.java.size" => collection::emit_size(chunks, current, line),
        "jvm.java.is_empty" => collection::emit_is_empty(chunks, current, line),
        "jvm.java.list_clear" => collection::emit_clear(chunks, current, line),
        "jvm.java.list_remove" => collection::emit_remove(chunks, current, argc, line),
        "jvm.java.list_remove_value" => collection::emit_remove(chunks, current, 3, line),
        "jvm.java.list_index_of" => collection::emit_index_of(chunks, current, line),
        "jvm.java.add_first" => collection::emit_add_first(chunks, current, line),
        "jvm.java.remove_first" => collection::emit_remove_first(chunks, current, line),
        "jvm.java.peek_first" => collection::emit_peek(chunks, current, false, line),
        "jvm.java.peek_last" => collection::emit_peek(chunks, current, true, line),
        "jvm.java.poll_first" => collection::emit_poll(chunks, current, false, line),
        "jvm.java.poll_last" => collection::emit_poll(chunks, current, true, line),
        "jvm.java.queue_poll" => collection::emit_poll(chunks, current, false, line),
        "jvm.java.priority_add" => collection::emit_priority_add(chunks, current, line),
        "jvm.java.priority_peek" => collection::emit_priority_peek(chunks, current, line),
        "jvm.java.sorted_add" => sorted_collection::emit_sorted_add(chunks, current, line),
        "jvm.java.sorted_first" => sorted_collection::emit_sorted_end(chunks, current, false, line),
        "jvm.java.sorted_last" => sorted_collection::emit_sorted_end(chunks, current, true, line),
        "jvm.java.sorted_ceiling" => collection::emit_sorted_bound(chunks, current, 0, line),
        "jvm.java.sorted_floor" => collection::emit_sorted_bound(chunks, current, 1, line),
        "jvm.java.sorted_higher" => collection::emit_sorted_bound(chunks, current, 2, line),
        "jvm.java.sorted_lower" => collection::emit_sorted_bound(chunks, current, 3, line),
        "jvm.java.sorted_descending_set" => {
            collection::emit_sorted_descending_set(chunks, current, line)
        }
        "jvm.java.hash_map_new" => map::emit_hash_map_new(chunks, current, argc, line),
        "jvm.java.concurrent_hash_map_new" => {
            map::emit_concurrent_hash_map_new(chunks, current, argc, line)
        }
        "jvm.java.identity_hash_map_new" => {
            map::emit_identity_hash_map_new(chunks, current, argc, line)
        }
        "jvm.java.linked_hash_map_new" => {
            map::emit_linked_hash_map_new(chunks, current, argc, line)
        }
        "jvm.java.map_of" => map::emit_map_of(chunks, current, argc, line),
        "jvm.java.map_entry" => map::emit_map_entry(chunks, current, line),
        "jvm.java.map_of_entries" => map::emit_map_of_entries(chunks, current, argc, line),
        "jvm.java.map_get" => map::emit_get(chunks, current, line),
        "jvm.java.map_put" => map::emit_put(chunks, current, line),
        "jvm.java.map_size" => map::emit_size(chunks, current, line),
        "jvm.java.map_is_empty" => map::emit_is_empty(chunks, current, line),
        "jvm.java.map_clear" => map::emit_clear(chunks, current, line),
        "jvm.java.map_key_set" => map::emit_key_set(chunks, current, line),
        "jvm.java.map_values" => map::emit_values(chunks, current, line),
        "jvm.java.entry_set" => map::emit_entry_set(chunks, current, line),
        "jvm.java.sorted_map_key_set" => map::emit_sorted_map_key_set(chunks, current, line),
        "jvm.java.sorted_map_first_key" => map::emit_sorted_map_key(chunks, current, false, line),
        "jvm.java.sorted_map_last_key" => map::emit_sorted_map_key(chunks, current, true, line),
        "jvm.java.sorted_map_higher_key" => {
            map::emit_sorted_map_bound_key(chunks, current, 2, line)
        }
        "jvm.java.sorted_map_lower_key" => map::emit_sorted_map_bound_key(chunks, current, 3, line),
        "jvm.java.sorted_map_values" => map::emit_sorted_map_values(chunks, current, line),
        "jvm.java.stringbuilder_new" => sb::emit_new(chunks, current, argc, line),
        "jvm.java.sb_length" => sb::emit_length(chunks, current, argc, line),
        "jvm.java.sb_is_empty" => sb::emit_is_empty(chunks, current, argc, line),
        "jvm.java.sb_is_not_empty" => sb::emit_is_not_empty(chunks, current, argc, line),
        "jvm.java.sb_insert" => sb::emit_insert(chunks, current, argc, line),
        "jvm.java.sb_delete" => sb::emit_delete(chunks, current, argc, line),
        "jvm.java.sb_reverse" => sb::emit_reverse(chunks, current, argc, line),
        "jvm.java.sb_set_length" => sb::emit_set_length(chunks, current, argc, line),
        "jvm.java.sb_clear" => sb::emit_clear(chunks, current, argc, line),
        "jvm.java.sb_set_char_at" => sb::emit_set_char_at(chunks, current, argc, line),
        "jvm.java.sb_char_at" => sb::emit_char_at(chunks, current, argc, line),
        "jvm.java.sb_append_code_point" => sb::emit_append_code_point(chunks, current, argc, line),
        "jvm.java.sb_append" => sb::emit_append(chunks, current, argc, line),
        "jvm.java.sb_append_line" => sb::emit_append_line(chunks, current, argc, line),
        "jvm.java.sb_to_string" => sb::emit_to_string(chunks, current, argc, line),
        "jvm.java.stringtokenizer_new" => st::emit_new(chunks, current, argc, line),
        "jvm.java.st_has_more" => st::emit_has_more(chunks, current, argc, line),
        "jvm.java.st_next" => st::emit_next(chunks, current, argc, line),
        "jvm.java.st_count" => st::emit_count(chunks, current, argc, line),

        _ => return false,
    }
    true
}
