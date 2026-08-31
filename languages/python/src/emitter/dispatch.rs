//! Auto-extracted `python.*` dispatch (language-specific routing lives in the
//! language module; the common dispatcher delegates here).

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    // Structural equality for composite built-ins — Python's `==` on two sets
    // is order-independent. Reached as `[builtin_slots.array] eq`; used to be
    // the `LanguageHooks::value_eq` callback, keyed by language name.
    if name == "python.value_eq" {
        crate::emitter::runtime_adapter::emit_py_value_eq(chunks, current, line);
        return true;
    }
    if let Some(exc_name) = name.strip_prefix("python.exc.") {
        crate::emitter::runtime_adapter::emit_py_exception(chunks, current, argc, exc_name, line);
        return true;
    }
    if let Some(exc_name) = name.strip_prefix("python.raise.") {
        crate::emitter::runtime_adapter::emit_py_raise(chunks, current, argc, exc_name, line);
        return true;
    }
    // Partly-implemented module surface. `argc == 0` is a bare `module.name`
    // read (the value), anything else is a call — see `surface_adapter`.
    if let Some(what) = name.strip_prefix("python.typeobj.") {
        crate::emitter::surface_adapter::emit_type_surface(chunks, current, argc, what, line);
        return true;
    }
    if let Some(what) = name.strip_prefix("python.surface.") {
        crate::emitter::surface_adapter::emit_function_surface(chunks, current, argc, what, line);
        return true;
    }
    match name {
        // hashlib / hmac over node:crypto — see hash_adapter.rs
        "python.hash_sha256" => {
            crate::emitter::hash_adapter::emit_sha256(chunks, current, argc, line)
        }
        "python.hash_sha512" => {
            crate::emitter::hash_adapter::emit_sha512(chunks, current, argc, line)
        }
        "python.hash_sha1" => crate::emitter::hash_adapter::emit_sha1(chunks, current, argc, line),
        "python.hash_md5" => crate::emitter::hash_adapter::emit_md5(chunks, current, argc, line),
        "python.hash_new" => crate::emitter::hash_adapter::emit_new(chunks, current, argc, line),
        "python.hash_hexdigest" => {
            crate::emitter::hash_adapter::emit_hexdigest(chunks, current, argc, line)
        }
        "python.hash_digest" => {
            crate::emitter::hash_adapter::emit_digest(chunks, current, argc, line)
        }
        "python.hmac_new" => {
            crate::emitter::hash_adapter::emit_hmac_new(chunks, current, argc, line)
        }
        "python.hmac_compare_digest" => {
            crate::emitter::hash_adapter::emit_compare_digest(chunks, current, argc, line)
        }
        "python.base64_b64encode" => {
            crate::emitter::base64_adapter::emit_b64encode(chunks, current, argc, line)
        }
        "python.base64_b64decode" => {
            crate::emitter::base64_adapter::emit_b64decode(chunks, current, argc, line)
        }
        "python.base64_urlsafe_b64encode" => {
            crate::emitter::base64_adapter::emit_urlsafe_b64encode(chunks, current, argc, line)
        }
        "python.base64_urlsafe_b64decode" => {
            crate::emitter::base64_adapter::emit_urlsafe_b64decode(chunks, current, argc, line)
        }
        "python.base64_encodebytes" => {
            crate::emitter::base64_adapter::emit_encodebytes(chunks, current, argc, line)
        }
        "python.base64_b16encode" => {
            crate::emitter::base64_adapter::emit_b16encode(chunks, current, argc, line)
        }
        "python.base64_b16decode" => {
            crate::emitter::base64_adapter::emit_b16decode(chunks, current, argc, line)
        }
        "python.base64_b32encode" => {
            crate::emitter::base64_adapter::emit_b32encode(chunks, current, argc, line)
        }
        "python.base64_b32decode" => {
            crate::emitter::base64_adapter::emit_b32decode(chunks, current, argc, line)
        }
        "python.base64_a85encode" => {
            crate::emitter::base64_adapter::emit_a85encode(chunks, current, argc, line)
        }
        "python.base64_a85decode" => {
            crate::emitter::base64_adapter::emit_a85decode(chunks, current, argc, line)
        }
        "python.base64_b85encode" => {
            crate::emitter::base64_adapter::emit_b85encode(chunks, current, argc, line)
        }
        "python.base64_b85decode" => {
            crate::emitter::base64_adapter::emit_b85decode(chunks, current, argc, line)
        }
        "python.binascii_b2a_base64" => {
            crate::emitter::base64_adapter::emit_b2a_base64(chunks, current, argc, line)
        }
        "python.binascii_a2b_base64" => {
            crate::emitter::base64_adapter::emit_a2b_base64(chunks, current, argc, line)
        }
        "python.binascii_hexlify" => {
            crate::emitter::base64_adapter::emit_hexlify(chunks, current, argc, line)
        }
        "python.binascii_unhexlify" => {
            crate::emitter::base64_adapter::emit_unhexlify(chunks, current, argc, line)
        }
        "python.binascii_crc32" => {
            crate::emitter::base64_adapter::emit_crc32(chunks, current, argc, line)
        }
        "python.zlib_adler32" => {
            crate::emitter::compression_adapter::emit_adler32(chunks, current, argc, line)
        }
        "python.glob_glob" => {
            crate::emitter::introspect_adapter::emit_glob(chunks, current, argc, line)
        }
        "python.quopri_encodestring" => {
            crate::emitter::quopri_locale_adapter::emit_encodestring(chunks, current, argc, line)
        }
        "python.quopri_decodestring" => {
            crate::emitter::quopri_locale_adapter::emit_decodestring(chunks, current, argc, line)
        }
        "python.locale_getlocale" => {
            crate::emitter::quopri_locale_adapter::emit_getlocale(chunks, current, argc, line)
        }
        "python.locale_getpreferredencoding" => {
            crate::emitter::quopri_locale_adapter::emit_getpreferredencoding(
                chunks, current, argc, line,
            )
        }
        "python.glob_escape" => {
            crate::emitter::introspect_adapter::emit_glob_escape(chunks, current, argc, line)
        }
        "python.linecache_getline" => {
            crate::emitter::introspect_adapter::emit_getline(chunks, current, argc, line)
        }
        "python.linecache_getlines" => {
            crate::emitter::introspect_adapter::emit_getlines(chunks, current, argc, line)
        }
        "python.linecache_none" => {
            crate::emitter::introspect_adapter::emit_linecache_none(chunks, current, argc, line)
        }
        "python.inspect_isclass" => {
            crate::emitter::introspect_adapter::emit_isclass(chunks, current, argc, line)
        }
        "python.inspect_isfunction" => {
            crate::emitter::introspect_adapter::emit_isfunction(chunks, current, argc, line)
        }
        "python.inspect_iscallable" => {
            crate::emitter::introspect_adapter::emit_iscallable(chunks, current, argc, line)
        }
        "python.inspect_getmembers" => {
            crate::emitter::introspect_adapter::emit_getmembers(chunks, current, argc, line)
        }
        "python.typing_cast" => {
            crate::emitter::typing_adapter::emit_cast(chunks, current, argc, line)
        }
        "python.typing_final" => {
            crate::emitter::typing_adapter::emit_marker_decorator(
                chunks, current, argc, "__final__", line,
            )
        }
        "python.typing_no_type_check" => {
            crate::emitter::typing_adapter::emit_marker_decorator(
                chunks,
                current,
                argc,
                "__no_type_check__",
                line,
            )
        }
        "python.typing_runtime_checkable" => {
            crate::emitter::typing_adapter::emit_marker_decorator(
                chunks,
                current,
                argc,
                "_is_runtime_protocol",
                line,
            )
        }
        "python.typing_typevar" => {
            crate::emitter::typing_adapter::emit_type_marker(chunks, current, argc, "TypeVar", line)
        }
        "python.typing_paramspec" => {
            crate::emitter::typing_adapter::emit_type_marker(
                chunks, current, argc, "ParamSpec", line,
            )
        }
        "python.typing_typevartuple" => {
            crate::emitter::typing_adapter::emit_type_marker(
                chunks,
                current,
                argc,
                "TypeVarTuple",
                line,
            )
        }
        "python.typing_newtype" => {
            crate::emitter::typing_adapter::emit_newtype(chunks, current, argc, line)
        }
        "python.typing_get_type_hints" => {
            crate::emitter::typing_adapter::emit_get_type_hints(chunks, current, argc, line)
        }
        "python.weakref_ref" => {
            crate::emitter::weakref_gc_adapter::emit_ref(chunks, current, argc, line)
        }
        "python.weakref_proxy" => {
            crate::emitter::weakref_gc_adapter::emit_proxy(chunks, current, argc, line)
        }
        "python.weakref_finalize" => {
            crate::emitter::weakref_gc_adapter::emit_finalize(chunks, current, argc, line)
        }
        "python.weakref_getweakrefcount" => {
            crate::emitter::weakref_gc_adapter::emit_getweakrefcount(chunks, current, argc, line)
        }
        "python.gc_collect" => {
            crate::emitter::weakref_gc_adapter::emit_gc_collect(chunks, current, argc, line)
        }
        "python.gc_get_count" => crate::emitter::weakref_gc_adapter::emit_gc_triple(
            chunks,
            current,
            argc,
            [0.0, 0.0, 0.0],
            line,
        ),
        "python.gc_get_threshold" => crate::emitter::weakref_gc_adapter::emit_gc_triple(
            chunks,
            current,
            argc,
            [700.0, 10.0, 10.0],
            line,
        ),
        "python.gc_get_stats" => {
            crate::emitter::weakref_gc_adapter::emit_gc_stats(chunks, current, argc, line)
        }
        "python.gc_is_tracked" => {
            crate::emitter::weakref_gc_adapter::emit_gc_is_tracked(chunks, current, argc, line)
        }
        "python.gc_zero" => {
            crate::emitter::weakref_gc_adapter::emit_gc_zero(chunks, current, argc, line)
        }
        "python.weakref_getweakrefs" => {
            crate::emitter::weakref_gc_adapter::emit_getweakrefs(chunks, current, argc, line)
        }
        "python.gc_empty_list" => {
            crate::emitter::weakref_gc_adapter::emit_gc_empty_list(chunks, current, argc, line)
        }
        "python.gc_isenabled" => {
            crate::emitter::weakref_gc_adapter::emit_gc_bool(chunks, current, argc, true, line)
        }
        "python.gc_none" => {
            crate::emitter::weakref_gc_adapter::emit_gc_none(chunks, current, argc, line)
        }
        "python.uuid4" => {
            crate::emitter::uuid_secrets_adapter::emit_uuid4(chunks, current, argc, line)
        }
        "python.uuid_new" => {
            crate::emitter::uuid_secrets_adapter::emit_uuid_new(chunks, current, argc, line)
        }
        "python.secrets_token_bytes" => {
            crate::emitter::uuid_secrets_adapter::emit_token_bytes(chunks, current, argc, line)
        }
        "python.secrets_token_hex" => {
            crate::emitter::uuid_secrets_adapter::emit_token_hex(chunks, current, argc, line)
        }
        "python.secrets_token_urlsafe" => {
            crate::emitter::uuid_secrets_adapter::emit_token_urlsafe(chunks, current, argc, line)
        }
        "python.secrets_randbelow" => {
            crate::emitter::uuid_secrets_adapter::emit_randbelow(chunks, current, argc, line)
        }
        "python.secrets_choice" => {
            crate::emitter::uuid_secrets_adapter::emit_choice(chunks, current, argc, line)
        }
        "python.zlib_compress" => {
            crate::emitter::compression_adapter::emit_zlib_compress(chunks, current, argc, line)
        }
        "python.zlib_decompress" => {
            crate::emitter::compression_adapter::emit_zlib_decompress(chunks, current, argc, line)
        }
        "python.gzip_compress" => {
            crate::emitter::compression_adapter::emit_gzip_compress(chunks, current, argc, line)
        }
        "python.gzip_decompress" => {
            crate::emitter::compression_adapter::emit_gzip_decompress(chunks, current, argc, line)
        }
        "python.colorsys_rgb_to_yiq" => {
            crate::emitter::colorsys_adapter::emit_rgb_to_yiq(chunks, current, argc, line)
        }
        "python.colorsys_yiq_to_rgb" => {
            crate::emitter::colorsys_adapter::emit_yiq_to_rgb(chunks, current, argc, line)
        }
        "python.colorsys_rgb_to_hls" => {
            crate::emitter::colorsys_adapter::emit_rgb_to_hls(chunks, current, argc, line)
        }
        "python.colorsys_hls_to_rgb" => {
            crate::emitter::colorsys_adapter::emit_hls_to_rgb(chunks, current, argc, line)
        }
        "python.colorsys_rgb_to_hsv" => {
            crate::emitter::colorsys_adapter::emit_rgb_to_hsv(chunks, current, argc, line)
        }
        "python.colorsys_hsv_to_rgb" => {
            crate::emitter::colorsys_adapter::emit_hsv_to_rgb(chunks, current, argc, line)
        }
        "python.codecs_encode" => {
            crate::emitter::base64_adapter::emit_codecs_encode(chunks, current, argc, line)
        }
        "python.codecs_decode" => {
            crate::emitter::base64_adapter::emit_codecs_decode(chunks, current, argc, line)
        }
        "python.codecs_lookup" => {
            crate::emitter::base64_adapter::emit_codecs_lookup(chunks, current, argc, line)
        }
        "python.codecs_escape_decode" => {
            crate::emitter::base64_adapter::emit_codecs_escape_decode(chunks, current, argc, line)
        }
        "python.codecs_escape_encode" => {
            crate::emitter::base64_adapter::emit_codecs_escape_encode(chunks, current, argc, line)
        }
        "python.codecs_iterencode" => {
            crate::emitter::base64_adapter::emit_codecs_iterencode(chunks, current, argc, line)
        }
        "python.codecs_iterdecode" => {
            crate::emitter::base64_adapter::emit_codecs_iterdecode(chunks, current, argc, line)
        }
        "python.unicodedata_normalize" => {
            crate::emitter::base64_adapter::emit_unicodedata_normalize(chunks, current, argc, line)
        }
        "python.first_arg" => {
            crate::emitter::base64_adapter::emit_first_arg(chunks, current, argc, line)
        }
        "python.json_dumps" => {
            crate::emitter::json_adapter::emit_json_dumps(chunks, current, argc, line);
        }
        "python.thread_start_with" => {
            crate::emitter::thread_adapter::emit_thread_start_with(chunks, current, line)
        }
        "python.thread_join" => {
            crate::emitter::thread_adapter::emit_thread_join(chunks, current, line)
        }
        "python.int_bit_length" => {
            crate::emitter::collections_adapter::emit_int_bit_length(chunks, current, argc, line)
        }
        "python.int_bit_count" => {
            crate::emitter::collections_adapter::emit_int_bit_count(chunks, current, argc, line)
        }
        "python.int_to_bytes" => {
            crate::emitter::collections_adapter::emit_int_to_bytes(chunks, current, argc, line)
        }
        "python.int_from_bytes" => {
            crate::emitter::collections_adapter::emit_int_from_bytes(chunks, current, argc, line)
        }
        "python.float_as_integer_ratio" => {
            crate::emitter::collections_adapter::emit_float_as_integer_ratio(
                chunks, current, argc, line,
            )
        }
        "python.enumerate" => {
            crate::emitter::collections_adapter::emit_enumerate(chunks, current, argc, line)
        }
        "python.tuple_from_iter" => {
            crate::emitter::collections_adapter::emit_tuple_from_iter(chunks, current, argc, line)
        }
        "python.sql_connect" => {
            crate::emitter::sql_adapter::emit_connect(chunks, current, argc, line)
        }
        "python.sql_cursor" => {
            crate::emitter::sql_adapter::emit_cursor(chunks, current, argc, line)
        }
        "python.sql_execute" => {
            crate::emitter::sql_adapter::emit_execute(chunks, current, argc, line)
        }
        "python.sql_executemany" => {
            crate::emitter::sql_adapter::emit_executemany(chunks, current, argc, line)
        }
        "python.sql_fetchall" => {
            crate::emitter::sql_adapter::emit_fetchall(chunks, current, argc, line)
        }
        "python.sql_fetchone" => {
            crate::emitter::sql_adapter::emit_fetchone(chunks, current, argc, line)
        }
        "python.sql_commit" => {
            crate::emitter::sql_adapter::emit_commit(chunks, current, argc, line)
        }
        "python.sql_rollback" => {
            crate::emitter::sql_adapter::emit_rollback(chunks, current, argc, line)
        }
        "python.sql_close" => crate::emitter::sql_adapter::emit_close(chunks, current, argc, line),
        "python.sql_begin" => crate::emitter::sql_adapter::emit_begin(chunks, current, argc, line),
        "python.math_factorial" => {
            crate::emitter::math_adapter::emit_factorial(chunks, current, argc, line)
        }
        "python.math_gcd" => crate::emitter::math_adapter::emit_gcd(chunks, current, argc, line),
        "python.math_lcm" => crate::emitter::math_adapter::emit_lcm(chunks, current, argc, line),
        "python.math_comb" => crate::emitter::math_adapter::emit_comb(chunks, current, argc, line),
        "python.math_perm" => crate::emitter::math_adapter::emit_perm(chunks, current, argc, line),
        "python.math_prod" => crate::emitter::math_adapter::emit_prod(chunks, current, argc, line),
        "python.math_degrees" => {
            crate::emitter::math_adapter::emit_degrees(chunks, current, argc, line)
        }
        "python.math_radians" => {
            crate::emitter::math_adapter::emit_radians(chunks, current, argc, line)
        }
        "python.math_copysign" => {
            crate::emitter::math_adapter::emit_copysign(chunks, current, argc, line)
        }
        "python.math_fmod" => crate::emitter::math_adapter::emit_fmod(chunks, current, argc, line),
        "python.math_ldexp" => {
            crate::emitter::math_adapter::emit_ldexp(chunks, current, argc, line)
        }
        "python.math_dist" => crate::emitter::math_adapter::emit_dist(chunks, current, argc, line),
        "python.math_modf" => crate::emitter::math_adapter::emit_modf(chunks, current, argc, line),
        "python.math_frexp" => {
            crate::emitter::math_adapter::emit_frexp(chunks, current, argc, line)
        }
        "python.math_isinf" => {
            crate::emitter::math_adapter::emit_isinf(chunks, current, argc, line)
        }
        "python.math_remainder" => {
            crate::emitter::math_adapter::emit_remainder(chunks, current, argc, line)
        }
        "python.math_isclose" => {
            crate::emitter::math_adapter::emit_isclose(chunks, current, argc, line)
        }
        "python.math_fsum" => crate::emitter::math_adapter::emit_fsum(chunks, current, argc, line),
        "python.extend" => crate::emitter::collections_adapter::emit_extend(chunks, current, line),
        "python.deque_extendleft" => {
            crate::emitter::collections_adapter::emit_extendleft(chunks, current, line)
        }
        "python.deque_rotate" => {
            crate::emitter::collections_adapter::emit_rotate(chunks, current, argc, line)
        }
        "python.move_to_end" => {
            crate::emitter::collections_adapter::emit_move_to_end(chunks, current, argc, line)
        }
        "python.popitem" => {
            crate::emitter::collections_adapter::emit_popitem(chunks, current, argc, line)
        }
        "python.counter_new" => {
            crate::emitter::collections_adapter::emit_counter_new(chunks, current, argc, line)
        }
        "python.get" => crate::emitter::collections_adapter::emit_get(chunks, current, argc, line),
        "python.pop" => crate::emitter::collections_adapter::emit_pop(chunks, current, argc, line),
        "python.str_find" => crate::emitter::string_adapter::emit_str_search(
            chunks, current, argc, line, false, false,
        ),
        "python.str_rfind" => crate::emitter::string_adapter::emit_str_search(
            chunks, current, argc, line, true, false,
        ),
        "python.str_rindex" => {
            crate::emitter::string_adapter::emit_str_search(chunks, current, argc, line, true, true)
        }
        "python.index" => {
            crate::emitter::collections_adapter::emit_index(chunks, current, argc, line)
        }
        "python.file_readline" => {
            crate::emitter::file_adapter::emit_readline(chunks, current, argc, line)
        }
        "python.file_writelines" => {
            crate::emitter::file_adapter::emit_writelines(chunks, current, argc, line)
        }
        "python.file_seek" => crate::emitter::file_adapter::emit_seek(chunks, current, argc, line),
        "python.file_tell" => crate::emitter::file_adapter::emit_tell(chunks, current, argc, line),
        "python.tmp_gettempdir" => {
            crate::emitter::file_adapter::emit_gettempdir(chunks, current, argc, line)
        }
        "python.tmp_mkdtemp" => {
            crate::emitter::file_adapter::emit_mkdtemp(chunks, current, argc, line)
        }
        "python.tmp_named" => {
            crate::emitter::file_adapter::emit_named_temp_file(chunks, current, argc, line)
        }
        "python.file_open" => crate::emitter::file_adapter::emit_open(chunks, current, argc, line),
        "python.file_read" => crate::emitter::file_adapter::emit_read(chunks, current, argc, line),
        "python.file_write" => {
            crate::emitter::file_adapter::emit_write(chunks, current, argc, line)
        }
        "python.file_readlines" => {
            crate::emitter::file_adapter::emit_readlines(chunks, current, argc, line)
        }
        "python.file_close" => {
            crate::emitter::file_adapter::emit_close(chunks, current, argc, line)
        }
        "python.shutil_copytree" => {
            crate::emitter::os_adapter::emit_copytree(chunks, current, argc, line)
        }
        "python.shutil_which" => {
            crate::emitter::os_adapter::emit_which(chunks, current, argc, line)
        }
        "python.tmp_mkstemp" => {
            crate::emitter::file_adapter::emit_mkstemp(chunks, current, argc, line)
        }
        "python.ospath_samefile" => {
            crate::emitter::file_adapter::emit_samefile(chunks, current, argc, line)
        }
        "python.tmp_prefix" => {
            crate::emitter::file_adapter::emit_tmp_prefix(chunks, current, argc, line)
        }
        "python.os_device_encoding" => {
            crate::emitter::os_adapter::emit_device_encoding(chunks, current, argc, line)
        }
        "python.os_term_size" => {
            crate::emitter::os_adapter::emit_term_size(chunks, current, argc, line)
        }
        "python.sys_getsizeof" => {
            crate::emitter::os_adapter::emit_getsizeof(chunks, current, argc, line)
        }
        "python.sys_intern" => crate::emitter::os_adapter::emit_intern(chunks, current, argc, line),
        "python.sys_getrecursionlimit" => {
            crate::emitter::os_adapter::emit_getrecursionlimit(chunks, current, argc, line)
        }
        "python.sys_setrecursionlimit" => {
            crate::emitter::os_adapter::emit_setrecursionlimit(chunks, current, argc, line)
        }
        "python.sys_encoding" => {
            crate::emitter::os_adapter::emit_encoding(chunks, current, argc, line)
        }
        "python.sys_is_finalizing" => {
            crate::emitter::os_adapter::emit_is_finalizing(chunks, current, argc, line)
        }
        "python.sys_exc_info" => {
            crate::emitter::os_adapter::emit_exc_info(chunks, current, argc, line)
        }
        "python.os_stat" => crate::emitter::os_adapter::emit_stat(chunks, current, argc, line),
        "python.os_entry_stat" => {
            crate::emitter::os_adapter::emit_entry_stat(chunks, current, argc, line)
        }
        "python.os_scandir" => {
            crate::emitter::os_adapter::emit_scandir(chunks, current, argc, line)
        }
        "python.os_walk" => crate::emitter::os_adapter::emit_walk(chunks, current, argc, line),
        "python.os_cpu_count" => {
            crate::emitter::os_adapter::emit_cpu_count(chunks, current, argc, line)
        }
        "python.os_fspath" => crate::emitter::os_adapter::emit_fspath(chunks, current, argc, line),
        "python.os_strerror" => {
            crate::emitter::os_adapter::emit_strerror(chunks, current, argc, line)
        }
        "python.os_is_file" => {
            crate::emitter::os_adapter::emit_entry_flag(chunks, current, "__is_file", line)
        }
        "python.os_is_dir" => {
            crate::emitter::os_adapter::emit_entry_flag(chunks, current, "__is_dir", line)
        }
        "python.os_is_symlink" => {
            crate::emitter::os_adapter::emit_entry_flag(chunks, current, "__is_link", line)
        }
        "python.os_inode" => crate::emitter::os_adapter::emit_entry_zero(chunks, current, line),
        "python.iter_array" => {
            crate::emitter::collections_adapter::emit_py_iter_array(chunks, current, argc, line)
        }
        "python.from_end" => {
            crate::emitter::collections_adapter::emit_from_end(chunks, current, argc, line)
        }
        "python.contains" => {
            crate::emitter::collections_adapter::emit_contains(chunks, current, line)
        }
        "python.attr_read" => {
            crate::emitter::collections_adapter::emit_attr_read(chunks, current, line)
        }
        "python.attr_write" => {
            crate::emitter::collections_adapter::emit_attr_write(chunks, current, line)
        }
        "python.attr_raw_write" => {
            crate::emitter::collections_adapter::emit_attr_raw_write(chunks, current, line)
        }
        "python.attr_delete" => {
            crate::emitter::collections_adapter::emit_attr_delete(chunks, current, line)
        }
        "python.getitem" => {
            crate::emitter::collections_adapter::emit_getitem(chunks, current, line)
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
        "python.it_tee" => crate::emitter::itertools_adapter::emit_tee(chunks, current, argc, line),
        "python.time_gmtime" => {
            crate::emitter::time_adapter::emit_gmtime(chunks, current, argc, line)
        }
        "python.time_struct_time" => {
            crate::emitter::time_adapter::emit_struct_time(chunks, current, argc, line)
        }
        "python.time_mktime" => {
            crate::emitter::time_adapter::emit_mktime(chunks, current, argc, line)
        }
        "python.time_asctime" => {
            crate::emitter::time_adapter::emit_asctime(chunks, current, argc, line)
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
        "python.astimezone" => {
            crate::emitter::datetime_adapter::emit_astimezone(chunks, current, argc, line)
        }
        "python.tzname" => {
            crate::emitter::datetime_adapter::emit_tzname(chunks, current, argc, line)
        }
        "python.tz_offset_str" => {
            crate::emitter::datetime_adapter::emit_tz_offset_str(chunks, current, argc, line)
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
        "python.dt_now" => crate::emitter::datetime_adapter::emit_now(chunks, current, argc, line),
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
        "python.dt_str" => crate::emitter::datetime_adapter::emit_str(chunks, current, argc, line),
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
        "python.gen_close" => {
            crate::emitter::collections_adapter::emit_gen_close(chunks, current, argc, line)
        }
        "python.frozenset" => {
            crate::emitter::collections_adapter::emit_frozenset(chunks, current, argc, line)
        }
        "python.frozenset_key" => {
            crate::emitter::collections_adapter::emit_frozenset_key(chunks, current, argc, line)
        }
        "python.sort_by_key" => {
            crate::emitter::collections_adapter::emit_sort_by_key(chunks, current, argc, line)
        }
        "python.min" => {
            crate::emitter::collections_adapter::emit_py_minmax(chunks, current, argc, false, line)
        }
        "python.max" => {
            crate::emitter::collections_adapter::emit_py_minmax(chunks, current, argc, true, line)
        }
        "python.sum" => {
            crate::emitter::collections_adapter::emit_py_sum(chunks, current, argc, line)
        }
        "python.reversed" => {
            crate::emitter::collections_adapter::emit_reversed(chunks, current, argc, line)
        }
        "python.iter_sentinel" => {
            crate::emitter::collections_adapter::emit_iter_sentinel(chunks, current, argc, line)
        }
        "python.zip_strict" => {
            crate::emitter::collections_adapter::emit_zip_strict(chunks, current, argc, line)
        }
        "python.zip_spread" => {
            crate::emitter::collections_adapter::emit_zip_spread(chunks, current, argc, line)
        }
        "python.dict_from_pairs" => {
            crate::emitter::collections_adapter::emit_dict_from_pairs(chunks, current, argc, line)
        }
        "python.re_search" => crate::emitter::re_adapter::emit_search(chunks, current, argc, line),
        "python.re_match" => crate::emitter::re_adapter::emit_match(chunks, current, argc, line),
        "python.re_findall" => {
            crate::emitter::re_adapter::emit_findall(chunks, current, argc, line)
        }
        "python.re_sub" => crate::emitter::re_adapter::emit_sub(chunks, current, argc, line),
        "python.re_split" => crate::emitter::re_adapter::emit_split(chunks, current, argc, line),
        "python.re_escape" => crate::emitter::re_adapter::emit_escape(chunks, current, argc, line),
        "python.make_set" => {
            crate::emitter::collections_adapter::emit_make_set(chunks, current, argc, line)
        }
        "python.set_issubset" => crate::emitter::collections_adapter::emit_set_predicate(
            chunks,
            current,
            "isSubsetOf",
            line,
        ),
        "python.set_issuperset" => crate::emitter::collections_adapter::emit_set_predicate(
            chunks,
            current,
            "isSupersetOf",
            line,
        ),
        "python.set_isdisjoint" => crate::emitter::collections_adapter::emit_set_predicate(
            chunks,
            current,
            "isDisjointFrom",
            line,
        ),
        "python.set_union" => {
            crate::emitter::collections_adapter::emit_set_union(chunks, current, argc, line)
        }
        "python.set_intersection" => {
            crate::emitter::collections_adapter::emit_set_intersection(chunks, current, argc, line)
        }
        "python.set_difference" => {
            crate::emitter::collections_adapter::emit_set_difference(chunks, current, argc, line)
        }
        "python.set_symmetric_difference" => {
            crate::emitter::collections_adapter::emit_set_symmetric_difference(
                chunks, current, argc, line,
            )
        }
        "python.add" => crate::emitter::collections_adapter::emit_add(chunks, current, line),
        "python.remove" => crate::emitter::collections_adapter::emit_remove(chunks, current, line),
        "python.discard" => {
            crate::emitter::collections_adapter::emit_discard(chunks, current, line)
        }
        "python.copy" => crate::emitter::collections_adapter::emit_copy(chunks, current, line),
        "python.update" => crate::emitter::collections_adapter::emit_update(chunks, current, line),
        "python.intersection_update" => {
            crate::emitter::collections_adapter::emit_intersection_update(chunks, current, line)
        }
        "python.difference_update" => {
            crate::emitter::collections_adapter::emit_difference_update(chunks, current, line)
        }
        "python.symmetric_difference_update" => {
            crate::emitter::collections_adapter::emit_symmetric_difference_update(
                chunks, current, line,
            )
        }
        "python.clear" => crate::emitter::collections_adapter::emit_clear(chunks, current, line),
        "python.length" => crate::emitter::collections_adapter::emit_length(chunks, current, line),
        "python.str_translate" => {
            crate::emitter::string_adapter::emit_translate(chunks, current, argc, line)
        }
        "python.str_maketrans" => {
            crate::emitter::string_adapter::emit_maketrans(chunks, current, argc, line)
        }
        "python.str_istitle" => {
            crate::emitter::string_adapter::emit_istitle(chunks, current, argc, line)
        }
        "python.str_casefold" => {
            crate::emitter::string_adapter::emit_casefold(chunks, current, argc, line)
        }
        "python.str_removeprefix" => {
            crate::emitter::string_adapter::emit_removeprefix(chunks, current, argc, line)
        }
        "python.str_removesuffix" => {
            crate::emitter::string_adapter::emit_removesuffix(chunks, current, argc, line)
        }
        "python.str_replace" => {
            crate::emitter::string_adapter::emit_replace(chunks, current, argc, line)
        }
        "python.str_startswith" => {
            crate::emitter::string_adapter::emit_startswith(chunks, current, argc, line)
        }
        "python.str_endswith" => {
            crate::emitter::string_adapter::emit_endswith(chunks, current, argc, line)
        }
        "python.str_count" => {
            crate::emitter::string_adapter::emit_count(chunks, current, argc, line)
        }
        "python.str_split" => {
            crate::emitter::string_adapter::emit_split(chunks, current, argc, line)
        }
        "python.str_rsplit" => {
            crate::emitter::string_adapter::emit_rsplit(chunks, current, argc, line)
        }
        "python.str_splitlines" => {
            crate::emitter::string_adapter::emit_splitlines(chunks, current, argc, line)
        }
        "python.str_expandtabs" => {
            crate::emitter::string_adapter::emit_expandtabs(chunks, current, argc, line)
        }
        "python.str" => crate::emitter::runtime_adapter::emit_str(chunks, current, argc, line),
        "python.repr" => crate::emitter::runtime_adapter::emit_repr(chunks, current, argc, line),
        "python.issubclass" => {
            crate::emitter::runtime_adapter::emit_issubclass(chunks, current, line)
        }
        "python.type" => crate::emitter::runtime_adapter::emit_py_type(chunks, current, line),
        "python.type_name" => {
            crate::emitter::runtime_adapter::emit_py_type_name(chunks, current, line)
        }
        "python.exception_instance" => {
            crate::emitter::runtime_adapter::emit_py_exception_instance(chunks, current, line)
        }
        "python.exception_message" => {
            crate::emitter::runtime_adapter::emit_py_exception_message(chunks, current, line)
        }
        "python.exception_add_note" => {
            crate::emitter::runtime_adapter::emit_py_exception_add_note(chunks, current, line)
        }
        "python.int" => crate::emitter::runtime_adapter::emit_py_int(chunks, current, argc, line),
        "python.ip4_parse" => {
            crate::emitter::socket_adapter::emit_ip4_parse(chunks, current, argc, line)
        }
        "python.ip4_str" => {
            crate::emitter::socket_adapter::emit_ip4_str(chunks, current, argc, line)
        }
        "python.ip4_octets" => {
            crate::emitter::socket_adapter::emit_ip4_octets(chunks, current, argc, line)
        }
        "python.ip4_mask" => {
            crate::emitter::socket_adapter::emit_ip4_mask(chunks, current, argc, line)
        }
        "python.ip4_count" => {
            crate::emitter::socket_adapter::emit_ip4_count(chunks, current, argc, line)
        }
        "python.ip4_net_parts" => {
            crate::emitter::socket_adapter::emit_ip4_net_parts(chunks, current, argc, line)
        }
        // The socket OBJECT. Was `VybeSocketImpl` in SOCKET_PRELUDE; it is
        // bytecode now so it cannot inherit walker defects the way python
        // source does. The walker rewrites `<handle>.m(...)` into
        // `__sock_m(<handle>, ...)` — see `rewrite_socket_call`.
        "python.sock_new" => crate::emitter::socket_adapter::emit_sock_new(chunks, current, argc, line),
        "python.sock_bind" => {
            crate::emitter::socket_adapter::emit_sock_bind(chunks, current, argc, line)
        }
        "python.sock_listen" => {
            crate::emitter::socket_adapter::emit_sock_listen(chunks, current, argc, line)
        }
        "python.sock_accept" => {
            crate::emitter::socket_adapter::emit_sock_accept(chunks, current, argc, line)
        }
        "python.sock_connect" => {
            crate::emitter::socket_adapter::emit_sock_connect(chunks, current, argc, line)
        }
        // `send` answers the byte count, `sendall` answers None. Same wire.
        "python.sock_send" => {
            crate::emitter::socket_adapter::emit_sock_send(chunks, current, argc, line, true)
        }
        "python.sock_sendall" => {
            crate::emitter::socket_adapter::emit_sock_send(chunks, current, argc, line, false)
        }
        "python.sock_recv" => {
            crate::emitter::socket_adapter::emit_sock_recv(chunks, current, argc, line)
        }
        "python.sock_getsockname" => {
            crate::emitter::socket_adapter::emit_sock_addr(chunks, current, argc, line, true)
        }
        "python.sock_getpeername" => {
            crate::emitter::socket_adapter::emit_sock_addr(chunks, current, argc, line, false)
        }
        "python.sock_close" => {
            crate::emitter::socket_adapter::emit_sock_close(chunks, current, argc, line)
        }
        "python.sock_settimeout" => {
            crate::emitter::socket_adapter::emit_sock_setopt(chunks, current, argc, "__timeout", line)
        }
        "python.sock_gettimeout" => {
            crate::emitter::socket_adapter::emit_sock_getopt(chunks, current, argc, "__timeout", line)
        }
        "python.sock_setsockopt" => {
            crate::emitter::socket_adapter::emit_sock_setopt(chunks, current, argc, "__sockopt", line)
        }
        "python.sock_getsockopt" => {
            crate::emitter::socket_adapter::emit_sock_getopt(chunks, current, argc, "__sockopt", line)
        }
        "python.sock_fileno" => {
            crate::emitter::socket_adapter::emit_sock_fileno(chunks, current, argc, line)
        }
        "python.sock_self" => {
            crate::emitter::socket_adapter::emit_sock_self(chunks, current, argc, line)
        }
        "python.sock_inet_aton" => {
            crate::emitter::socket_adapter::emit_inet_aton(chunks, current, argc, line)
        }
        "python.sock_inet_ntoa" => {
            crate::emitter::socket_adapter::emit_inet_ntoa(chunks, current, argc, line)
        }
        "python.sock_getservbyname" => {
            crate::emitter::socket_adapter::emit_getservbyname(chunks, current, argc, line)
        }
        "python.sock_gethostname" => {
            crate::emitter::socket_adapter::emit_gethostname(chunks, current, argc, line)
        }
        "python.sock_gethostbyname" => {
            crate::emitter::socket_adapter::emit_gethostbyname(chunks, current, argc, line)
        }
        "python.sock_getaddrinfo" => {
            crate::emitter::socket_adapter::emit_getaddrinfo(chunks, current, argc, line)
        }
        "python.url_join" => crate::emitter::url_adapter::emit_urljoin(chunks, current, argc, line),
        "python.url_split" => {
            crate::emitter::url_adapter::emit_urlsplit(chunks, current, argc, line)
        }
        "python.url_unsplit" => {
            crate::emitter::url_adapter::emit_urlunsplit(chunks, current, argc, line)
        }
        "python.url_encode" => {
            crate::emitter::url_adapter::emit_urlencode(chunks, current, argc, line)
        }
        "python.url_parse_qs" => {
            crate::emitter::url_adapter::emit_parse_qs(chunks, current, argc, line)
        }
        "python.url_parse_qsl" => {
            crate::emitter::url_adapter::emit_parse_qsl(chunks, current, argc, line)
        }
        "python.url_quote" => crate::emitter::url_adapter::emit_quote(chunks, current, argc, line),
        "python.url_quote_plus" => {
            crate::emitter::url_adapter::emit_quote_plus(chunks, current, argc, line)
        }
        "python.url_unquote" => {
            crate::emitter::url_adapter::emit_unquote(chunks, current, argc, line)
        }
        "python.url_unquote_plus" => {
            crate::emitter::url_adapter::emit_unquote_plus(chunks, current, argc, line)
        }
        "python.lock_acquire" => {
            crate::emitter::lock_adapter::emit_lock_acquire(chunks, current, argc, line)
        }
        "python.lock_release" => {
            crate::emitter::lock_adapter::emit_lock_release(chunks, current, argc, line)
        }
        "python.calendar_new" => {
            crate::emitter::calendar_adapter::emit_calendar_new(chunks, current, argc, line)
        }
        "python.calendar_html_new" => crate::emitter::calendar_adapter::emit_calendar_typed_new(
            chunks,
            current,
            argc,
            "HTMLCalendar",
            line,
        ),
        "python.calendar_text_new" => crate::emitter::calendar_adapter::emit_calendar_typed_new(
            chunks,
            current,
            argc,
            "TextCalendar",
            line,
        ),
        "python.calendar_leapdays" => {
            crate::emitter::calendar_adapter::emit_leapdays(chunks, current, argc, line)
        }
        "python.calendar_timegm" => {
            crate::emitter::calendar_adapter::emit_timegm(chunks, current, argc, line)
        }
        "python.calendar_monthcalendar" => {
            crate::emitter::calendar_adapter::emit_monthcalendar(chunks, current, argc, line)
        }
        "python.calendar_itermonthdays" => {
            crate::emitter::calendar_adapter::emit_itermonthdays(chunks, current, argc, line)
        }
        "python.calendar_itermonthdays2" => {
            crate::emitter::calendar_adapter::emit_itermonthdays2(chunks, current, argc, line)
        }
        "python.calendar_yeardayscalendar" => {
            crate::emitter::calendar_adapter::emit_yeardayscalendar(chunks, current, argc, line)
        }
        "python.calendar_text_formatmonth" => {
            crate::emitter::calendar_adapter::emit_text_formatmonth(chunks, current, argc, line)
        }
        "python.calendar_html_formatmonth" => {
            crate::emitter::calendar_adapter::emit_html_formatmonth(chunks, current, argc, line)
        }
        "python.calendar_setfirstweekday" => {
            crate::emitter::calendar_adapter::emit_setfirstweekday(chunks, current, argc, line)
        }
        "python.calendar_firstweekday" => {
            crate::emitter::calendar_adapter::emit_firstweekday(chunks, current, argc, line)
        }
        "python.vars" => crate::emitter::runtime_adapter::emit_vars(chunks, current, argc, line),
        "python.is_dataclass" => {
            crate::emitter::dataclass_adapter::emit_is_dataclass(chunks, current, argc, line)
        }
        "python.dataclass_asdict" => {
            crate::emitter::dataclass_adapter::emit_asdict(chunks, current, argc, line)
        }
        "python.dataclass_astuple" => {
            crate::emitter::dataclass_adapter::emit_astuple(chunks, current, argc, line)
        }
        "python.dataclass_fields" => {
            crate::emitter::dataclass_adapter::emit_fields(chunks, current, argc, line)
        }
        "python.dir" => crate::emitter::runtime_adapter::emit_dir(chunks, current, argc, line),
        "python.hasattr" => crate::emitter::runtime_adapter::emit_hasattr(chunks, current, line),
        "python.getattr" => {
            crate::emitter::runtime_adapter::emit_getattr(chunks, current, argc, line)
        }
        "python.setattr" => crate::emitter::runtime_adapter::emit_setattr(chunks, current, line),
        "python.delattr" => crate::emitter::runtime_adapter::emit_delattr(chunks, current, line),
        "python.print" => crate::emitter::runtime_adapter::emit_print(chunks, current, argc, line),
        "python.bytes_decode" => {
            crate::emitter::runtime_adapter::emit_bytes_decode(chunks, current, argc, line)
        }
        "python.struct_pack" => {
            crate::emitter::struct_adapter::emit_struct_pack(chunks, current, argc, line)
        }
        "python.struct_unpack" => {
            crate::emitter::struct_adapter::emit_struct_unpack(chunks, current, argc, line)
        }
        "python.struct_calcsize" => {
            crate::emitter::struct_adapter::emit_struct_calcsize(chunks, current, argc, line)
        }
        "python.struct_unpack_from" => {
            crate::emitter::struct_adapter::emit_struct_unpack_from(chunks, current, argc, line)
        }
        "python.struct_pack_into" => {
            crate::emitter::struct_adapter::emit_struct_pack_into(chunks, current, argc, line)
        }
        "python.struct_iter_unpack" => {
            crate::emitter::struct_adapter::emit_struct_iter_unpack(chunks, current, argc, line)
        }
        "python.struct_new" => {
            crate::emitter::struct_adapter::emit_struct_new(chunks, current, argc, line)
        }
        "python.pyneg" => crate::emitter::runtime_adapter::emit_pyneg(chunks, current, line),
        "python.pylt" => crate::emitter::runtime_adapter::emit_pylt(chunks, current, line),
        "python.pygt" => crate::emitter::runtime_adapter::emit_pygt(chunks, current, line),
        "python.pyle" => crate::emitter::runtime_adapter::emit_pyle(chunks, current, line),
        "python.pyge" => crate::emitter::runtime_adapter::emit_pyge(chunks, current, line),
        "python.pyadd" => crate::emitter::runtime_adapter::emit_pyadd(chunks, current, line),
        "python.pymul" => crate::emitter::runtime_adapter::emit_pymul(chunks, current, line),
        "python.pysub" => crate::emitter::runtime_adapter::emit_pysub(chunks, current, line),
        "python.pytruediv" => {
            crate::emitter::runtime_adapter::emit_pytruediv(chunks, current, line)
        }
        "python.pyfloordiv" => {
            crate::emitter::runtime_adapter::emit_pyfloordiv(chunks, current, line)
        }
        "python.pymod" => crate::emitter::runtime_adapter::emit_pymod(chunks, current, line),
        "python.fmt_fixed" => {
            crate::emitter::runtime_adapter::emit_py_fmt_fixed(chunks, current, line)
        }
        "python.fmt_sci" => crate::emitter::runtime_adapter::emit_py_fmt_sci(chunks, current, line),
        "python.fmt_group" => {
            crate::emitter::runtime_adapter::emit_py_fmt_group(chunks, current, line)
        }
        "python.pypow" => crate::emitter::runtime_adapter::emit_pypow(chunks, current, line),
        "python.range" => crate::emitter::runtime_adapter::emit_range(chunks, current, argc, line),
        "python.ospath_join" => {
            crate::emitter::os_path_adapter::emit_join(chunks, current, argc, line)
        }
        "python.ospath_split" => {
            crate::emitter::os_path_adapter::emit_split(chunks, current, argc, line)
        }
        "python.ospath_splitext" => {
            crate::emitter::os_path_adapter::emit_splitext(chunks, current, argc, line)
        }
        "python.ospath_basename" => {
            crate::emitter::os_path_adapter::emit_basename(chunks, current, argc, line)
        }
        "python.ospath_dirname" => {
            crate::emitter::os_path_adapter::emit_dirname(chunks, current, argc, line)
        }
        "python.ospath_normpath" => {
            crate::emitter::os_path_adapter::emit_normpath(chunks, current, argc, line)
        }
        "python.ospath_realpath" => {
            crate::emitter::os_path_adapter::emit_realpath(chunks, current, argc, line)
        }
        "python.ospath_abspath" => {
            crate::emitter::os_path_adapter::emit_realpath(chunks, current, argc, line)
        }
        "python.ospath_isabs" => {
            crate::emitter::os_path_adapter::emit_isabs(chunks, current, argc, line)
        }
        "python.ospath_normcase" => {
            crate::emitter::os_path_adapter::emit_normcase(chunks, current, argc, line)
        }
        "python.ospath_expandvars" => {
            crate::emitter::os_path_adapter::emit_expandvars(chunks, current, argc, line)
        }
        "python.ospath_expanduser" => {
            crate::emitter::os_path_adapter::emit_expanduser(chunks, current, argc, line)
        }
        "python.ospath_islink" => {
            crate::emitter::os_path_adapter::emit_islink(chunks, current, argc, line)
        }
        "python.ospath_ismount" => {
            crate::emitter::os_path_adapter::emit_ismount(chunks, current, argc, line)
        }
        "python.ospath_relpath" => {
            crate::emitter::os_path_adapter::emit_relpath(chunks, current, argc, line)
        }
        "python.ospath_commonprefix" => {
            crate::emitter::os_path_adapter::emit_commonprefix(chunks, current, argc, line)
        }
        "python.ospath_commonpath" => {
            crate::emitter::os_path_adapter::emit_commonpath(chunks, current, argc, line)
        }
        name if crate::emitter::runtime_adapter::emit_helper(name, chunks, current, argc, line) => {
        }

        // ── COBOL adapters ──
        _ => return false,
    }
    true
}
