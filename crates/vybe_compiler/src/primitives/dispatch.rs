//! Centralized `common:<name>` emit dispatcher.
//!
//! Language profiles use the `emit = "common:<category>.<op>"` convention to
//! delegate to a canonical compiler_common helper. This module owns the
//! `<name> → emit fn` mapping so every language compiler shares one source
//! of truth — adding a new common op only needs to be done here, and every
//! frontend that uses the dispatcher gets it for free.
//!
//! ## Two flavors
//!
//! - `emit_common(name, chunk, line)` handles ops that need ONLY a chunk and
//!   line (the vast majority — pure bytecode emits).
//! - `emit_common_with_imports(name, chunk, line, import)` handles ops that
//!   ALSO need to register a host import (e.g. `threading.sleep` adds a
//!   `wasi:clocks::sleep` import). The `import` callback resolves the import
//!   index in whatever way the host compiler does (typically by adding to a
//!   designated chunk's import table).
//!
//! Both functions return `true` if they recognized and emitted `name`, and
//! `false` if the name is unknown — letting the caller fall through to its
//! own dispatch for language-specific common ops.

use vybe_runtime::Chunk;
use vybe_ast::{BitLane, FloatLane, MidpointPolicy, NumericRepr};
use vybe_runtime::opcode::Op;

use crate::primitives::threading as thread_adapter;
use crate::primitives::{
    base64, collections, config, csv, dict, fs_path, heap, http_cookie, http_form,
    http_request_env, http_session,
    io, object, ops, paths, reflection, sets, strings, threading, url, xml,
};

/// Handle common ops that need only a chunk and line.
/// Returns `true` if `name` was recognized and emitted, `false` otherwise.
///
/// `argc` is the number of caller-supplied values currently on the
/// stack at the emit site. Most emits ignore it (their stack contract
/// is fixed — `dict.has` always pops two), but multi-arity emits like
/// .NET constructors with overloaded shapes (`new StringBuilder()` vs
/// `new StringBuilder("initial")`) branch on it to pick the right
/// bytecode.
///
/// Takes `&mut Vec<Chunk>` rather than `&mut [Chunk]` because some helpers
/// (e.g. `threading.task_delay`) push a new function chunk for the worker
/// body. Slice-shape ops still work — `&mut Vec<Chunk>` derefs to
/// `&mut [Chunk]` for index access.
pub fn emit_common(
    name: &str,
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) -> bool {
    // Language-specific routing lives in each language's emitter module,
    // registered the same way file extensions are: a language sets its
    // `emit_dispatch` in the `languages::all()` registry (and shared
    // platforms like `dotnet` register via `emitter::platform_emit_dispatch`).
    // The common dispatcher here only owns the genuinely-shared
    // `common:<cat>.*` keys below (collections/dict/strings/threading/…).
    // Adding a language never touches this file.
    if let Some(dot) = name.find('.') {
        if let Some(dispatch) = crate::languages::emit_dispatch_for(&name[..dot]) {
            if dispatch(name, chunks, current, argc, line) {
                return true;
            }
        }
    }
    match name {
        // ── Request environment ──
        // The request as language-neutral data, read from `wasi:http`. PHP's
        // `$_SERVER`/`$_GET`, WSGI `environ` and Rack `env` are all renames of
        // these; see `documentation/httpserver.md` §4a.
        "http_request.method" => http_request_env::emit_method(chunks, current, line),
        "http_request.path" => http_request_env::emit_path(chunks, current, line),
        "http_request.path_with_query" => {
            http_request_env::emit_path_with_query(chunks, current, line)
        }
        "http_request.query_string" => http_request_env::emit_query_string(chunks, current, line),
        "http_request.scheme" => http_request_env::emit_scheme(chunks, current, line),
        "http_request.authority" => http_request_env::emit_authority(chunks, current, line),
        "http_request.headers" => http_request_env::emit_headers(chunks, current, line),
        "http_request.environ" => http_request_env::emit_environ(chunks, current, line),
        // Cookies are not in any spec surface — see `primitives/http_cookie`.
        "http_cookie.serialize" => http_cookie::emit_serialize(chunks, current, argc, line),
        "http_cookie.request_cookies" => http_cookie::emit_request_cookies(chunks, current, line),
        "http_form.parsed_body" => http_form::emit_parsed_body(chunks, current, line),
        "http_form.fields" => http_form::emit_body_fields(chunks, current, line),
        "http_form.files" => http_form::emit_body_files(chunks, current, line),
        "http_session.id" => http_session::emit_id(chunks, current, "SESSIONID", line),
        "http_session.new_id" => http_session::emit_new_id(chunks, current, line),
        "http_session.set_id" => http_session::emit_set_id(chunks, current, line),
        "http_session.name" => http_session::emit_name(chunks, current, "SESSIONID", line),
        "http_session.set_name" => http_session::emit_set_name(chunks, current, line),
        "http_session.status" => http_session::emit_status(chunks, current, line),
        "http_session.data" => http_session::emit_data(chunks, current, line),
        "http_session.start" => http_session::emit_start(chunks, current, "SESSIONID", line),
        "http_session.regenerate_id" => http_session::emit_regenerate_id(chunks, current, line),
        "http_session.destroy" => http_session::emit_destroy(chunks, current, line),
        "http_session.unset" => http_session::emit_unset(chunks, current, line),
        "http_request.body" => http_request_env::emit_body(chunks, current, line),
        "http_request.request_params" => {
            http_request_env::emit_request_params(chunks, current, line)
        }
        "http_request.query_params" => http_request_env::emit_query_params(chunks, current, line),
        "http_form.parse_multipart" => http_form::emit_parse_multipart(chunks, current, line),
        "http_form.parse_urlencoded" => http_form::emit_parse_urlencoded(chunks, current, line),
        "http_cookie.parse" => http_cookie::emit_parse_cookie_header(chunks, current, line),

        // ── Output buffering ──
        // Language-neutral on purpose: `ob_*` is PHP's SPELLING of a capability
        // (capture what would have been written), not a PHP feature. A language
        // maps its own names onto these in its profile; the semantics live in
        // `io.rs` beside the write they intercept.
        // ── Filesystem ──
        // Path-addressed file operations, lowered onto `wasi:filesystem@0.3.1`
        // in `fs_path.rs`. Language-neutral because a path is: PHP's
        // `file_get_contents`, Python's `open().read()`, Ruby's `File.read`
        // and Pascal's `AssignFile` are spellings of the same capability.
        //
        // These exist so a profile can bind to the SPEC lowering by name.
        // Before them a profile could only say `host:wasi:filesystem:readFile`
        // — a verb that is not in the WIT — because a host import is the only
        // thing profile syntax could reach. That is precisely why thirty
        // invented verbs accumulated inside a real WASI namespace: the honest
        // call was a sequence (preopens → open-at → read-via-stream → drain)
        // and nothing but a shim could be named in one line.
        // ── Path strings ──
        // NOT filesystem: `Path.GetExtension("a/b.txt")` is ".txt" whether or
        // not the file exists. WASI has no path interface because path SYNTAX
        // is a language concern, which is why these ten sat in
        // `wasi:filesystem` naming functions no WIT declares, to do work that
        // never needed a host call. Two of them do read a capability and say so
        // at their definition.
        "path.file_name" => paths::emit_file_name(&mut chunks[current], line),
        "path.directory" => paths::emit_directory(&mut chunks[current], line),
        "path.extension" => paths::emit_extension(&mut chunks[current], line),
        "path.file_stem" => paths::emit_file_stem(&mut chunks[current], line),
        "path.has_extension" => paths::emit_has_extension(&mut chunks[current], line),
        "path.change_extension" => paths::emit_change_extension(&mut chunks[current], line),
        "path.is_rooted" => paths::emit_is_rooted(&mut chunks[current], line),
        "path.combine" => paths::emit_combine(&mut chunks[current], argc, line),
        "path.full_path" => paths::emit_full_path(&mut chunks[current], line),
        "path.temp_path" => paths::emit_temp_path(&mut chunks[current], line),

        // ── Component Model streams ──
        // What replaced `wasi:io`. 0.3.1 deleted that package outright: a
        // stream is a Component Model TYPE now, so reading and writing one is
        // `canon stream.{read,write}` — canonical built-ins the compiler
        // emits, not host imports. A profile row can only name a host import,
        // so without these two names a language had nothing to call, which is
        // the whole reason `wasi:io/streams:read` survived in two profiles
        // long after the package it names ceased to exist.
        //
        // Language-neutral for the same reason `filesystem.*` above is: a
        // `stream<u8>` parameter is a WIT type, not a feature of C or Python.
        // Sinks and sources that take or return one stay ordinary rows.
        "stream.read_bytes" => io::emit_read_stream_chunk(&mut chunks[current], line),
        "stream.drain_bytes" => io::emit_read_stream_to_bytes(&mut chunks[current], line),
        "stream.from_bytes" => io::emit_bytes_to_stream(&mut chunks[current], line),
        "stream.read_handle" => io::emit_read_stream_handle(&mut chunks[current], line),
        "stream.try_read_handle" => io::emit_try_read_stream_handle(&mut chunks[current], line),

        "filesystem.read_file" => fs_path::emit_read_file(&mut chunks[current], line),
        "filesystem.read_file_bytes" => fs_path::emit_read_file_bytes(&mut chunks[current], line),
        "filesystem.write_file" => fs_path::emit_write_file(&mut chunks[current], line),
        "filesystem.append_file" => fs_path::emit_append_file(&mut chunks[current], line),
        "filesystem.exists" => fs_path::emit_exists(&mut chunks[current], line),
        "filesystem.is_file" => fs_path::emit_is_file(&mut chunks[current], line),
        "filesystem.is_dir" => fs_path::emit_is_dir(&mut chunks[current], line),
        "filesystem.file_size" => fs_path::emit_file_size(&mut chunks[current], line),
        "filesystem.stat" => fs_path::emit_stat(&mut chunks[current], line),
        "filesystem.mkdir" => fs_path::emit_mkdir(&mut chunks[current], line),
        "filesystem.mkdir_all" => fs_path::emit_mkdir_all(chunks, current, line),
        "filesystem.unlink" => fs_path::emit_unlink(&mut chunks[current], line),
        "filesystem.rmdir" => fs_path::emit_rmdir(&mut chunks[current], line),
        "filesystem.remove" => fs_path::emit_remove(&mut chunks[current], line),
        "filesystem.remove_all" => fs_path::emit_remove_all(&mut chunks[current], line),
        "filesystem.rename" => fs_path::emit_rename(&mut chunks[current], line),
        "filesystem.copy" => fs_path::emit_copy(&mut chunks[current], line),
        // ── Numbered file handles ──
        //
        // `Open #1 For Output`, Pascal's `AssignFile`/`Rewrite`, and every other
        // language whose file API is a NUMBER rather than an object. The handle
        // table is a guest global mapping file number → `{path, mode, pos}`;
        // WASI has no cursor, because `read-via-stream`/`write-via-stream` take
        // a `filesize` and are deliberately stateless.
        //
        // These had no `common:` name until now, so the only way to reach them
        // was a Rust call from an emitter — which is why Pascal's profile still
        // named `host:wasi:filesystem:openFile`, a verb no WIT declares, while
        // the real lowering sat one crate away.
        "filesystem.open_file" => fs_path::emit_open_file(chunks, current, argc, line),
        "filesystem.close_file" => fs_path::emit_close_file(chunks, current, line),
        "filesystem.print_file" => fs_path::emit_print_file(chunks, current, argc, line),
        "filesystem.write_file_handle" => {
            fs_path::emit_write_file_handle(chunks, current, argc, line)
        }
        "filesystem.line_input" => fs_path::emit_line_input(chunks, current, line),
        "filesystem.input_file" => fs_path::emit_input_file(chunks, current, line),

        // Enumeration. `list_dir` answers names, `read_dir_entries` answers the
        // WIT's `{ type, name }` record — a language that wants `isFile` asks
        // the second and compares `type` itself, because `isFile` is that
        // language's question and `descriptor-type` is the spec's answer.
        "filesystem.list_dir" => {
            // `os.listdir()` and `Dir.entries` take the path as OPTIONAL and
            // mean the working directory when it is absent. WASI has no
            // implicit cwd — `open-at` needs something to resolve — so the
            // default is spelled here rather than left to each caller.
            if argc == 0 {
                chunks[current].emit_string_const(".", line);
            }
            fs_path::emit_list_dir(&mut chunks[current], line)
        }
        "filesystem.read_dir_entries" => {
            fs_path::emit_read_directory_entries(&mut chunks[current], line)
        }

        "output_buffer.start" => io::emit_ob_start(chunks, current, argc, line),
        "output_buffer.get_level" => io::emit_ob_get_level(chunks, current, line),
        "output_buffer.get_contents" => io::emit_ob_get_contents(chunks, current, line),
        "output_buffer.get_length" => io::emit_ob_get_length(chunks, current, None, line),
        "output_buffer.clean" => io::emit_ob_clean(chunks, current, line),
        "output_buffer.end_clean" => io::emit_ob_end_clean(chunks, current, line),
        "output_buffer.end_flush" => io::emit_ob_end_flush(chunks, current, line),
        "output_buffer.get_clean" => io::emit_ob_get_clean(chunks, current, line),
        "output_buffer.get_flush" => io::emit_ob_get_flush(chunks, current, line),
        "output_buffer.flush" => io::emit_ob_flush(chunks, current, line),

        // ── Dict ops ──
        "dict.set_dynamic" => {
            dict::emit_set_dynamic(chunks, current, line);
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            // void return
        }
        "dict.get_dynamic" => dict::emit_get_dynamic(chunks, current, line),
        "dict.has" => dict::emit_method_has(chunks, current, line),
        "dict.delete" => dict::emit_method_delete(chunks, current, line),
        "dict.clear" => dict::emit_method_clear_stack(chunks, current, line),
        "dict.size" => dict::emit_method_size(chunks, current, line),
        "dict.keys" => dict::emit_keys(chunks, current, line),
        "dict.values" => dict::emit_values(chunks, current, line),
        "dict.items" => dict::emit_items(chunks, current, line),
        "dict.new" => dict::emit_new(chunks, current, line),

        // ── Set ops ──
        // Language adapters normalize their surface (`setOf`, Python set
        // literals, Pascal set operators, .NET HashSet) to these operations.
        // The storage is ECMA Set, but that stays behind this primitive layer.
        "sets.new" => sets::emit_new(chunks, current, line),
        "sets.literal" => sets::emit_literal(chunks, current, argc, line),
        "sets.from_iterable" => sets::emit_from_iterable(chunks, current, line),
        "sets.add" => sets::emit_add(chunks, current, line),
        "sets.add_changed" => sets::emit_add_changed(chunks, current, line),
        "sets.delete" => sets::emit_delete(chunks, current, line),
        "sets.has" => sets::emit_has(chunks, current, line),
        "sets.size" => sets::emit_size(chunks, current, line),
        "sets.clear" => sets::emit_clear(chunks, current, line),
        "sets.union" => sets::emit_union(chunks, current, line),
        "sets.union_with" => sets::emit_union_with(chunks, current, line),
        "sets.intersection" => sets::emit_intersection(chunks, current, line),
        "sets.intersect_with" => sets::emit_intersect_with(chunks, current, line),
        "sets.difference" => sets::emit_difference(chunks, current, line),
        "sets.except_with" => sets::emit_except_with(chunks, current, line),
        "sets.symmetric_difference" => sets::emit_symmetric_difference(chunks, current, line),
        "sets.symmetric_except_with" => sets::emit_symmetric_except_with(chunks, current, line),
        "sets.is_subset_of" => sets::emit_is_subset_of(chunks, current, line),
        "sets.is_superset_of" => sets::emit_is_superset_of(chunks, current, line),
        "sets.is_disjoint_from" => sets::emit_is_disjoint_from(chunks, current, line),
        "sets.values_array" => sets::emit_values_array(chunks, current, line),

        // ── Object ops ── ecma:object/new creates a plain JS Object.
        // The `.NET` Dictionary class uses this as its backing (matches
        // the ECMA-262 rule that a Dictionary<string, T> is shape-identical
        // to an Object). Method dispatch routes through `ecma:object/*`
        // via TypeRegistry, so no parallel vybe:types/dict* host fns are
        // consulted.
        "object.new" => {
            // Import MUST be registered on the chunk that emits the call —
            // registering on chunks[0] gives an index that is out of range of
            // the current chunk's table, and the normalize pass's script-table
            // fallback maps it correctly only by luck (mis-resolved to
            // js-string.concat in nested function-expression contexts).
            let idx = chunks[current].add_import("ecma:object", "new");
            chunks[current].emit_call(idx, 0, line);
        }
        "object.equals" => object::emit_equals(&mut chunks[current], line),
        "object.is_null" => object::emit_is_null(&mut chunks[current], line),
        "object.non_null" => object::emit_non_null(&mut chunks[current], line),
        "object.hash_code" => object::emit_hash_code(&mut chunks[current], line),
        "object.hash" => {
            collections::emit_array_new(chunks, current, argc as u16, line);
            object::emit_hash_array(&mut chunks[current], line);
        }
        "object.hash_array" => object::emit_hash_array(&mut chunks[current], line),
        "object.compare" => {
            let abi = crate::primitives::class_context::module_receiver_abi(chunks);
            object::emit_compare(&mut chunks[current], abi, line)
        }
        "object.to_string_or" => {
            if argc < 2 {
                chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            }
            object::emit_to_string_or(&mut chunks[current], line);
        }
        // Text conversion that resolves the ToString ROLE first and falls back
        // to the ECMA `String()` coercion. A language whose print path binds
        // this reaches a user's string conversion whatever the CLASS's own
        // language spelled it — Go's `String()`, Python's `__str__`, Ruby's
        // `to_s` — because the lookup is the numeric slot, not a name.
        "object.to_string_role" => {
            let slot = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
            crate::primitives::expressions::emit_rich_to_string(&mut chunks[current], slot, line);
        }

        // ── Reflection substrate ──
        // These are runtime primitives only. Language adapters retain their
        // surface quirks (`Reflect.*`, PHP ReflectionClass, Go reflect.Value),
        // but all three can share the same object/property/type operations.
        "reflection.typeof" => reflection::emit_typeof(chunks, current, line),
        "reflection.is_callable" => reflection::emit_is_callable(chunks, current, line),
        "reflection.get" => {
            reflection::emit_reflect_op(chunks, current, reflection::ReflectOp::Get, argc, line)
        }
        "reflection.set" => {
            reflection::emit_reflect_op(chunks, current, reflection::ReflectOp::Set, argc, line)
        }
        "reflection.has_own" => reflection::emit_has_own(chunks, current, line),
        "reflection.has_in" => reflection::emit_has_in(chunks, current, line),
        "reflection.keys" => {
            reflection::emit_object_view(chunks, current, reflection::ObjectKeysMode::Own, line)
        }
        "reflection.for_in" => {
            reflection::emit_object_view(chunks, current, reflection::ObjectKeysMode::ForIn, line)
        }
        "reflection.values" => {
            reflection::emit_object_view(chunks, current, reflection::ObjectKeysMode::Values, line)
        }
        "reflection.entries" => {
            reflection::emit_object_view(chunks, current, reflection::ObjectKeysMode::Entries, line)
        }
        "reflection.instanceof" => reflection::emit_instanceof(chunks, current, line),
        "reflection.object" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Object, argc, line)
        }
        "reflection.assign" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Assign, argc, line)
        }
        "reflection.freeze" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Freeze, argc, line)
        }
        "reflection.from_entries" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::FromEntries,
            argc,
            line,
        ),
        "reflection.create" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Create, argc, line)
        }
        "reflection.seal" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Seal, argc, line)
        }
        "reflection.is_frozen" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::IsFrozen, argc, line)
        }
        "reflection.is_sealed" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::IsSealed, argc, line)
        }
        "reflection.object_is" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Is, argc, line)
        }
        "reflection.get_prototype_of" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::GetPrototypeOf,
            argc,
            line,
        ),
        "reflection.get_own_property_names" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::GetOwnPropertyNames,
            argc,
            line,
        ),
        "reflection.get_own_property_descriptor" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::GetOwnPropertyDescriptor,
            argc,
            line,
        ),
        "reflection.get_own_property_descriptors" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::GetOwnPropertyDescriptors,
            argc,
            line,
        ),
        "reflection.get_own_property_symbols" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::GetOwnPropertySymbols,
            argc,
            line,
        ),
        "reflection.define_property" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::DefineProperty,
            argc,
            line,
        ),
        "reflection.define_properties" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::DefineProperties,
            argc,
            line,
        ),
        "reflection.prevent_extensions" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::PreventExtensions,
            argc,
            line,
        ),
        "reflection.is_extensible" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::IsExtensible,
            argc,
            line,
        ),
        "reflection.set_prototype_of" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::SetPrototypeOf,
            argc,
            line,
        ),
        "reflection.group_by" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::GroupBy, argc, line)
        }
        "reflection.delete" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Delete, argc, line)
        }
        "reflection.object_get" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Get, argc, line)
        }
        "reflection.object_set" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::Set, argc, line)
        }
        "reflection.track_key" => {
            reflection::emit_object_op(chunks, current, reflection::ObjectOp::TrackKey, argc, line)
        }
        "reflection.property_is_enumerable" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::PropertyIsEnumerable,
            argc,
            line,
        ),
        "reflection.has_own_property" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::HasOwnProperty,
            argc,
            line,
        ),
        "reflection.is_prototype_of" => reflection::emit_object_op(
            chunks,
            current,
            reflection::ObjectOp::IsPrototypeOf,
            argc,
            line,
        ),
        "reflection.apply" => {
            reflection::emit_reflect_op(chunks, current, reflection::ReflectOp::Apply, argc, line)
        }
        "reflection.construct" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::Construct,
            argc,
            line,
        ),
        "reflection.delete_property" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::DeleteProperty,
            argc,
            line,
        ),
        "reflection.reflect_define_property" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::DefineProperty,
            argc,
            line,
        ),
        "reflection.reflect_get_own_property_descriptor" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::GetOwnPropertyDescriptor,
            argc,
            line,
        ),
        "reflection.reflect_get_prototype_of" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::GetPrototypeOf,
            argc,
            line,
        ),
        "reflection.reflect_has" => {
            reflection::emit_reflect_op(chunks, current, reflection::ReflectOp::Has, argc, line)
        }
        "reflection.reflect_is_extensible" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::IsExtensible,
            argc,
            line,
        ),
        "reflection.own_keys" => {
            reflection::emit_reflect_op(chunks, current, reflection::ReflectOp::OwnKeys, argc, line)
        }
        "reflection.reflect_prevent_extensions" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::PreventExtensions,
            argc,
            line,
        ),
        "reflection.reflect_set_prototype_of" => reflection::emit_reflect_op(
            chunks,
            current,
            reflection::ReflectOp::SetPrototypeOf,
            argc,
            line,
        ),

        // ── XML qualified names ──
        // Portable QName shape shared by Go xml.Name, .NET XName, Java QName,
        // DOM nodes, and Vybe XML values. Full XML parsing/tree work stays in
        // the host `web:dom-parser`; these are only language-surface adapters.
        "xml.name" => xml::emit_name(chunks, current, argc, line),
        "xml.local" => xml::emit_local(chunks, current, argc, line),
        "xml.namespace" => xml::emit_namespace(chunks, current, argc, line),
        "xml.prefix" => xml::emit_prefix(chunks, current, argc, line),
        "xml.qualified" => xml::emit_qualified(chunks, current, argc, line),
        "xml.equal" => xml::emit_equal(chunks, current, argc, line),
        "xml.from_dom_node" => xml::emit_from_dom_node(chunks, current, argc, line),
        "xml.node_name" => xml::emit_node_name(chunks, current, argc, line),
        "xml.parse" => xml::emit_parse(chunks, current, argc, line),
        "xml.load" => xml::emit_load(chunks, current, argc, line),
        "xml.save" => xml::emit_save(chunks, current, argc, line),
        "xml.elements" => xml::emit_elements(chunks, current, argc, line),
        "xml.attribute" => xml::emit_attribute(chunks, current, argc, line),

        // ── Collection ops (route through ecma:array imports; the helper
        // registers on chunks[current] — the chunk it emits into — for the
        // reason spelled out in the `object.new` arm above). ──
        "collections.push" => collections::emit_push(chunks, current, line),
        "collections.pop" => collections::emit_pop(chunks, current, line),
        "collections.length" => collections::emit_len(chunks, current, line),
        "collections.rank" => collections::emit_rank(chunks, current, line),
        "collections.get" => collections::emit_get(chunks, current, line),
        "collections.set" => collections::emit_set(chunks, current, line),
        "collections.contains" => collections::emit_contains(chunks, current, line),
        "tuple.value_eq" => {
            crate::primitives::tuples::emit_tuple_value_eq(&mut chunks[current], line)
        }
        // Floored modulo — the result takes the DIVISOR's sign, so `-7 % 3` is
        // 2 and not -1. Python's `%` and Dart's both work this way; C's and
        // JS's `%` truncate instead. Reached as the `Mod` slot binding.
        "math.floor_mod" => {
            crate::primitives::math::emit_python_floor_mod(&mut chunks[current], line)
        }
        "collections.index_of" => collections::emit_index_of(chunks, current, line),
        "collections.last_index_of" => collections::emit_last_index_of(chunks, current, line),
        "collections.remove_at" => collections::emit_remove_at(chunks, current, line),
        "collections.sorted" => collections::emit_sorted(chunks, current, line),
        "collections.reverse" => collections::emit_reverse(chunks, current, line),
        "collections.join" => collections::emit_join(chunks, current, line),
        "collections.join_sep_first" => collections::emit_join_sep_first(chunks, current, line),
        "collections.slice" => collections::emit_slice(chunks, current, line),
        "collections.slice_with_bound" => collections::emit_slice_with_bound(chunks, current, line),
        "collections.new" => collections::emit_array_new(chunks, current, 0, line),
        "collections.shift" => collections::emit_shift(chunks, current, line),
        "collections.concat" => collections::emit_concat(chunks, current, line),
        "collections.fill" => collections::emit_fill(chunks, current, line),
        "collections.fill_all" => collections::emit_fill_all(chunks, current, line),
        "collections.copy_to" => collections::emit_copy_to(chunks, current, line),
        "collections.repeat_value" => collections::emit_repeat_value(chunks, current, line),
        "collections.sort" => collections::emit_sort(chunks, current, line),
        "collections.identity" => collections::emit_identity(chunks, current, line),
        "collections.first_of_two" => collections::emit_first_of_two(chunks, current, line),
        "collections.nil_to_empty" => collections::emit_nil_to_empty_array(chunks, current, line),
        "collections.clear_keyed" => collections::emit_clear_keyed(chunks, current, line),
        "collections.index_func" => collections::emit_index_func(chunks, current, line),
        "collections.sort_func" => collections::emit_sort_func(chunks, current, line),
        "collections.is_sorted_func" => collections::emit_is_sorted_func(chunks, current, line),
        "collections.binary_search_pair" => {
            collections::emit_binary_search_pair(chunks, current, line)
        }
        "collections.binary_search_func_pair" => {
            collections::emit_binary_search_func_pair(chunks, current, line)
        }
        "collections.index_of_from" => collections::emit_index_of_from(chunks, current, line),
        "collections.last_index_of_from" => {
            collections::emit_last_index_of_from(chunks, current, line)
        }
        "collections.remove_range" => collections::emit_remove_range(chunks, current, line),
        "collections.get_range" => collections::emit_get_range(chunks, current, line),
        "collections.clone" => collections::emit_clone(chunks, current, line),
        "collections.sequence_equal" => collections::emit_sequence_equal(chunks, current, line),
        "collections.delete_range_copy" => {
            collections::emit_delete_range_copy(chunks, current, line)
        }
        "collections.insert_range_copy" => {
            collections::emit_insert_range_copy(chunks, current, line)
        }
        "collections.replace_range_copy" => {
            collections::emit_replace_range_copy(chunks, current, line)
        }
        "collections.compact_adjacent" => collections::emit_compact_adjacent(chunks, current, line),
        "collections.map_clone" => collections::emit_map_clone(chunks, current, line),
        "collections.map_copy" => collections::emit_map_copy(chunks, current, line),
        "collections.map_delete_func" => collections::emit_map_delete_func(chunks, current, line),
        "collections.sequence_compare" => collections::emit_sequence_compare(chunks, current, line),
        "collections.is_sorted" => collections::emit_is_sorted(chunks, current, line),
        "collections.insert_range" => collections::emit_insert_range(chunks, current, line),
        "collections.set_range" => collections::emit_set_range(chunks, current, line),
        "collections.binary_search" => collections::emit_binary_search(chunks, current, line),
        "collections.reverse_range" => collections::emit_reverse_range(chunks, current, line),
        "collections.remove" => collections::emit_remove_value(chunks, current, line),
        "collections.insert" => collections::emit_insert_at(chunks, current, line),
        "collections.clear" => collections::emit_clear(chunks, current, line),
        "collections.sum" => collections::emit_sum(chunks, current, line),
        "collections.min" => collections::emit_pymin(chunks, current, line),
        "collections.max" => collections::emit_pymax(chunks, current, line),
        // Range materialisation (Kotlin/JVM style)
        // `collections.range_inc`  – 2-arg inclusive:  [start..=end]  step +1
        "collections.range_inc" => collections::emit_range(chunks, current, 2, true, line),
        // `collections.range_desc` – 2-arg descending: (start downTo end) step -1
        // Emitted as reversed(end..=start) to avoid duplicating loop logic.
        "collections.range_desc" => {
            // Stack: [start, end].  swap → [end, start], emit inclusive range,
            // then reverse in-place so we get [start, start-1, …, end].
            let tmp = chunks[current].alloc_scratch(2);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, tmp, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_SET, tmp + 1, line);
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, tmp, line); // end
            chunks[current].emit_op_u16(vybe_runtime::opcode::Op::LOCAL_GET, tmp + 1, line); // start
            collections::emit_range(chunks, current, 2, true, line); // [end..=start]
            collections::emit_reverse(chunks, current, line); // reversed → [start..end]
        }
        // `collections.range_step` – 3-arg strided: (start, stop_exclusive, step)
        "collections.range_step" => collections::emit_range(chunks, current, 3, false, line),

        // ── Shared heap / priority-queue primitives ──
        "heap.heapify" => heap::emit_heapify(chunks, current, argc, line),
        "heap.push" => heap::emit_push(chunks, current, argc, line),
        "heap.pop" => heap::emit_pop(chunks, current, argc, line),
        "heap.replace" => heap::emit_replace(chunks, current, argc, line),
        "heap.push_pop" => heap::emit_push_pop(chunks, current, argc, line),
        "heap.nsmallest" => heap::emit_nsmallest(chunks, current, argc, line),
        "heap.nlargest" => heap::emit_nlargest(chunks, current, argc, line),
        "heap.merge" => heap::emit_merge(chunks, current, argc, line),

        // ── Python adapters ──
        "strings.join_iterable" => strings::emit_join_iterable(chunks, current, line),
        "strings.length" => strings::emit_length(&mut chunks[current], line),
        "strings.char_code" => strings::emit_char_code(&mut chunks[current], line),
        "strings.from_char_code" => strings::emit_from_char_code(&mut chunks[current], line),
        "strings.to_upper" => strings::emit_to_upper(&mut chunks[current], line),
        "strings.to_lower" => strings::emit_to_lower(&mut chunks[current], line),
        "strings.trim" => strings::emit_trim(&mut chunks[current], line),
        "strings.substring" => strings::emit_substring(&mut chunks[current], line),
        "strings.replace" => strings::emit_replace(&mut chunks[current], line),
        "strings.split" => strings::emit_split(&mut chunks[current], line),
        "strings.index_of" => strings::emit_index_of(&mut chunks[current], line),
        "strings.concat" => strings::emit_concat(&mut chunks[current], 2, line),
        "base64.encode_binary_string" => base64::emit_encode_binary_string(chunks, current, line),
        "base64.decode_binary_string" => base64::emit_decode_binary_string(chunks, current, line),
        "sprintf.format" => crate::primitives::sprintf::emit_sprintf(chunks, current, argc, line),
        "sprintf.format_array" => {
            crate::primitives::sprintf::emit_sprintf_from_array(chunks, current, line)
        }

        // ── Expression ops ──
        "expressions.undefined" => {
            crate::primitives::expressions::emit_undefined(&mut chunks[current], line)
        }
        "expressions.i32_not" => {
            crate::primitives::expressions::emit_i32_not(&mut chunks[current], line)
        }
        "expressions.f64_mod" => {
            crate::primitives::expressions::emit_f64_mod(&mut chunks[current], line)
        }
        "math.clamp" => crate::primitives::math::emit_clamp(&mut chunks[current], line),
        "math.copysign" => crate::primitives::math::emit_copysign(&mut chunks[current], line),
        "math.signbit" => crate::primitives::math::emit_signbit(&mut chunks[current], line),
        "math.dim" => crate::primitives::math::emit_dim(&mut chunks[current], line),
        "math.nan" => crate::primitives::math::emit_nan(&mut chunks[current], line),
        "math.inf" => crate::primitives::math::emit_inf(&mut chunks[current], line),
        "math.is_inf" => crate::primitives::math::emit_is_inf(&mut chunks[current], line),
        // Reinterpretation — a REPRESENTATION operation, so it lives with the
        // bit family rather than with IEEE arithmetic. The `math.*` spellings
        // just below are Go's, kept working and pointed at this same code.
        "bits.reinterpret_i32" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::I32, line)
        }
        "bits.reinterpret_f32" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::F32, line)
        }
        "bits.reinterpret_i64" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::I64, line)
        }
        "bits.reinterpret_f64" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::F64, line)
        }
        // Next representable float. One home for `nextafter` / `nextUp` /
        // `nextDown` / `ulp` — C, Java, Kotlin and Fortran each had their own
        // and all of them were wrong.
        // Round-to-integral, one row per midpoint policy. Languages disagree
        // three ways and each row says which one it means.
        "math.min_num" => crate::primitives::math::emit_min_num(&mut chunks[current], line),
        "math.max_num" => crate::primitives::math::emit_max_num(&mut chunks[current], line),
        "math.round_half_even" => {
            crate::primitives::math::emit_round(&mut chunks[current], MidpointPolicy::HalfEven, line)
        }
        "math.round_half_away" => crate::primitives::math::emit_round(
            &mut chunks[current],
            MidpointPolicy::HalfAwayFromZero,
            line,
        ),
        "math.round_half_up" => {
            crate::primitives::math::emit_round(&mut chunks[current], MidpointPolicy::HalfUp, line)
        }
        "math.next_up32" => {
            crate::primitives::math::emit_next_toward(&mut chunks[current], true, FloatLane::F32, line)
        }
        "math.next_up64" => {
            crate::primitives::math::emit_next_toward(&mut chunks[current], true, FloatLane::F64, line)
        }
        "math.next_down32" => {
            crate::primitives::math::emit_next_toward(&mut chunks[current], false, FloatLane::F32, line)
        }
        "math.next_down64" => {
            crate::primitives::math::emit_next_toward(&mut chunks[current], false, FloatLane::F64, line)
        }
        "math.next_after32" => {
            crate::primitives::math::emit_next_after(&mut chunks[current], FloatLane::F32, line)
        }
        "math.next_after64" => {
            crate::primitives::math::emit_next_after(&mut chunks[current], FloatLane::F64, line)
        }
        "math.ulp32" => {
            crate::primitives::math::emit_ulp(&mut chunks[current], FloatLane::F32, line)
        }
        "math.ulp64" => {
            crate::primitives::math::emit_ulp(&mut chunks[current], FloatLane::F64, line)
        }
        "math.f64_bits" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::I64, line)
        }
        "math.f64_from_bits" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::F64, line)
        }
        // The bit family. A bare-name spelling (`POPCNT`, `bits.OnesCount`) must
        // stay SHADOWABLE by a user definition, which is what profile
        // resolution gives it — so it routes here rather than being folded in a
        // walker. The AST operator nodes call the same functions, so there is
        // one implementation with two entry points.
        // Elementwise over a boolean SEQUENCE — a bit set. See `bits.rs`.
        "bits.log2" => crate::primitives::bits::emit_log2(chunks, current, line),
        // The UNSIGNED 32-bit view of an i32 result — `F64_CONVERT_I32_U`.
        // A language whose declared type is unsigned needs this after any
        // bitwise op, whose lane is signed: `Not 0UI` is `4294967295`, not −1.
        "bits.as_unsigned32" => {
            crate::primitives::bits::emit_as_unsigned32(&mut chunks[current], line)
        }
        "bits.rotl32u" => crate::primitives::bits::emit_rotl32_unsigned(chunks, current, line),
        "bits.rotr32u" => crate::primitives::bits::emit_rotr32_unsigned(chunks, current, line),
        "bits.is_pow2" => crate::primitives::bits::emit_is_pow2(chunks, current, line),
        "bits.round_up_pow2" => crate::primitives::bits::emit_round_up_pow2(chunks, current, line),
        "bits.array_and" => crate::primitives::bits::emit_array_and(chunks, current, line),
        "bits.array_or" => crate::primitives::bits::emit_array_or(chunks, current, line),
        "bits.array_xor" => crate::primitives::bits::emit_array_xor(chunks, current, line),
        "bits.array_not" => crate::primitives::bits::emit_array_not(chunks, current, line),
        "bits.popcnt32" => {
            crate::primitives::bits::emit_pop_count(&mut chunks[current], BitLane::W32, line)
        }
        "bits.popcnt64" => {
            crate::primitives::bits::emit_pop_count(&mut chunks[current], BitLane::W64, line)
        }
        "bits.clz32" => {
            crate::primitives::bits::emit_leading_zeros(&mut chunks[current], BitLane::W32, line)
        }
        "bits.clz64" => {
            crate::primitives::bits::emit_leading_zeros(&mut chunks[current], BitLane::W64, line)
        }
        "bits.ctz32" => {
            crate::primitives::bits::emit_trailing_zeros(&mut chunks[current], BitLane::W32, line)
        }
        "bits.ctz64" => {
            crate::primitives::bits::emit_trailing_zeros(&mut chunks[current], BitLane::W64, line)
        }
        "bits.rotl32" => {
            crate::primitives::bits::emit_rotate(&mut chunks[current], BitLane::W32, true, line)
        }
        "bits.rotl64" => {
            crate::primitives::bits::emit_rotate(&mut chunks[current], BitLane::W64, true, line)
        }
        "bits.rotr32" => {
            crate::primitives::bits::emit_rotate(&mut chunks[current], BitLane::W32, false, line)
        }
        "bits.rotr64" => {
            crate::primitives::bits::emit_rotate(&mut chunks[current], BitLane::W64, false, line)
        }
        "math.f32_bits" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::I32, line)
        }
        "math.f32_from_bits" => {
            crate::primitives::bits::emit_reinterpret(&mut chunks[current], NumericRepr::F32, line)
        }
        "expressions.bool_not" => {
            crate::primitives::expressions::emit_bool_not(&mut chunks[current], line)
        }

        // ── Delegate ops ──
        "delegates.combine" => crate::primitives::delegates::emit_combine(chunks, current, line),
        "delegates.remove" => crate::primitives::delegates::emit_remove(chunks, current, line),
        "delegates.invocation_list" => {
            crate::primitives::delegates::emit_get_invocation_list(chunks, current, line)
        }

        // ── JS Node compatibility adapters ───────────────────────
        // JS source keeps the Node-shaped call surface, but lowers
        // through compile-time adapters that compose the real
        // `wasi:sockets/*` interfaces. These live in the shared
        // `.NET` adapter home under `platforms/dotnet/core` so every
        // frontend can reuse them without a JS-only emitter fork.
        "threading.task_run" => threading::emit_task_run(chunks, current, line),
        "threading.task_delay" => threading::emit_task_delay(chunks, current, line),
        "threading.thread_new" => threading::emit_thread_new(chunks, current, line),
        "threading.thread_start" => threading::emit_thread_start(&mut chunks[current], line),
        "threading.thread_join" => threading::emit_thread_join(&mut chunks[current], line),
        "threading.thread_spawn" => threading::emit_thread_spawn(chunks, current, line),
        "threading.monitor_notify" | "threading.monitor_notify_all" => {
            object::emit_monitor_notify(&mut chunks[current], line)
        }

        // ── VB Choose / Switch — variadic 1-indexed selector ────────
        // `Choose(idx, v1, v2, ..., vN)` returns `vidx`. Variadic so it
        // needs `argc` threading; can't be a stdlib chunk (fixed arity).
        // Implementation: pack trailing vals into an array via
        // `ARRAY_NEW_FIXED`, save to a local, then `ARRAY_GET array[idx-1]`.
        // .NET-shape rather than stdlib because Choose is a VB.NET / VBA
        // language built-in, not a generic helper.
        "threading.atomic_load" => threading::emit_atomic_load(&mut chunks[current], line),
        "threading.atomic_store" => threading::emit_atomic_store(&mut chunks[current], line),
        "threading.atomic_add" => threading::emit_atomic_add(&mut chunks[current], line),
        "threading.atomic_sub" => threading::emit_atomic_sub(&mut chunks[current], line),
        "threading.atomic_xchg" => threading::emit_atomic_xchg(&mut chunks[current], line),
        "threading.atomic_cmpxchg" => threading::emit_atomic_cmpxchg(&mut chunks[current], line),
        "threading.atomic_fence" => threading::emit_atomic_fence(&mut chunks[current], line),
        "threading.suspend" => threading::emit_suspend(&mut chunks[current], line),

        // ── String ops (profile common:str_*) ──
        "str_reverse" => strings::emit_str_reverse(&mut chunks[current], line),
        "str_length" => strings::emit_length(&mut chunks[current], line),
        // `scalar` unit (Unicode code points) — the counterparts to the UTF-16
        // `str_*` arms above. A language binds these where its surface counts
        // code points (PHP `mb_*`, Python `str`). See strings.rs.
        "str_scalar_length" => strings::emit_scalar_length(chunks, current, line),
        "str_scalar_substring" => strings::emit_scalar_substring(chunks, current, line),
        "str_scalar_index_of" => strings::emit_scalar_index_of(chunks, current, line),
        "str_scalar_chars" => strings::emit_scalar_chars(&mut chunks[current], line),
        // `byte` unit (UTF-8 octets) — php `strlen`, Lua `#`, Go `len(s)`. This
        // is a BINDABLE target on purpose: `unifiedstringplan.md` §3a settled
        // that the index unit is the VALUE of a `[builtin_slots.string] len`
        // binding, not a parameter threaded through the emitters, so a language
        // that counts bytes declares it rather than being special-cased. Until
        // this arm existed there was nothing on the platform to point at — the
        // only byte counter was private to php's adapter.
        "str_byte_length" => strings::emit_byte_length(chunks, current, line),
        // Character-class predicates — tier-3 adapter primitives (no ECMA-262
        // string surface defines them). The `(String, Is*)` platform-default
        // slot rows point here; the receiver must already BE a string, so a
        // platform with a non-string char model (JVM lone-surrogate numbers)
        // guards at its own call site. See strings.rs section note.
        "str_is_digit" => strings::emit_is_digit(chunks, current, line),
        "str_is_alpha" => strings::emit_is_alpha(chunks, current, line),
        "str_is_alnum" => strings::emit_is_alnum(chunks, current, line),
        "str_is_space" => strings::emit_is_space(chunks, current, line),
        "str_is_upper" => strings::emit_is_upper(chunks, current, line),
        "str_is_lower" => strings::emit_is_lower(chunks, current, line),
        "str_cstr_length" => strings::emit_cstr_length(chunks, current, line),
        "str_cstr_truncate" => strings::emit_cstr_truncate(chunks, current, line),
        // Shell-style glob matching — php `fnmatch`, python `fnmatch.fnmatch`,
        // Go `path.Match`, Ruby `File.fnmatch`. Stack: `[name, pattern]`.
        // `_fold` is the case-insensitive spelling; it is the regex `i` flag,
        // not a lower-casing of both sides, so `[A-Z]` keeps its meaning.
        // Thousands grouping over a formatted numeric string — php
        // `number_format`, python's `,` format flag, java `%,d`.
        // Stack: `[formatted, group_sep, dec_point]`.
        "str_group_digits" => strings::emit_group_digits(chunks, current, line),
        // `[s, index, count, insert]` → spliced string; index is 0-based.
        "str_splice" => strings::emit_splice(chunks, current, line),
        // CSV — a structured format, so it lives beside `json` and `url`.
        // Dialect (delimiter, enclosure) is on the STACK because php takes it
        // as runtime arguments.
        // INI/config text — the same argument as CSV one entry down, and
        // the same runtime-dialect rule: key CASE is a stack value because
        // python lowercases option names and php does not.
        "config.parse" => config::emit_parse(chunks, current, line),
        "csv.parse_line" => csv::emit_parse_line(chunks, current, line),
        // A whole document, because a newline inside an enclosure is CONTENT and
        // no pre-split can know that.
        "csv.parse_document" => csv::emit_parse_document(chunks, current, line),
        "csv.format_row" => {
            csv::emit_format_row(chunks, current, csv::FormatOptions::minimal(), line)
        }
        // fpc `TStringList.CommaText` also encloses on whitespace.
        "csv.format_row_quote_ws" => csv::emit_format_row(
            chunks,
            current,
            csv::FormatOptions::quote_whitespace(),
            line,
        ),
        "str_glob_match" => {
            strings::emit_glob_match(chunks, current, strings::GlobOptions::exact(), line)
        }
        "str_glob_match_fold" => {
            strings::emit_glob_match(chunks, current, strings::GlobOptions::folded(), line)
        }
        // Charlist trim — the adapter primitive. `ecma:string.trim` takes no
        // character set, so php/python/ruby each grew a copy. Languages whose
        // default set IS the ECMA one bind the `_ws` forms, which delegate.
        // URL percent-encoding, reachable from ANY profile so a language
        // BINDS the shared codec instead of reimplementing it. The four
        // variants are measured: php `urlencode` == java `URLEncoder`
        // (`form`), go `QueryEscape` == python `quote_plus`
        // (`form_rfc3986`), php `rawurlencode`/.NET `EscapeDataString`
        // (`rfc3986`), python `quote` (`path`).
        // Canonical component reads, receiver on the stack — what a
        // PROFILE builtin gets. java `getProtocol`, php `parse_url` and
        // python `urlsplit` all read the same nine components.
        "url.component_scheme" => {
            url::emit_component_of(chunks, current, url::UrlField::Scheme, line)
        }
        "url.component_user" => url::emit_component_of(chunks, current, url::UrlField::User, line),
        "url.component_pass" => url::emit_component_of(chunks, current, url::UrlField::Pass, line),
        "url.component_host" => url::emit_component_of(chunks, current, url::UrlField::Host, line),
        "url.component_port" => url::emit_component_of(chunks, current, url::UrlField::Port, line),
        "url.component_netloc" => {
            url::emit_component_of(chunks, current, url::UrlField::Netloc, line)
        }
        "url.component_path" => url::emit_component_of(chunks, current, url::UrlField::Path, line),
        "url.component_query" => {
            url::emit_component_of(chunks, current, url::UrlField::Query, line)
        }
        "url.component_fragment" => {
            url::emit_component_of(chunks, current, url::UrlField::Fragment, line)
        }
        "url.encode_form" => {
            url::emit_percent_encode(chunks, current, url::PercentOptions::form(), line)
        }
        "url.encode_form_rfc3986" => {
            url::emit_percent_encode(chunks, current, url::PercentOptions::form_rfc3986(), line)
        }
        "url.encode_rfc3986" => {
            url::emit_percent_encode(chunks, current, url::PercentOptions::rfc3986(), line)
        }
        "url.encode_path" => {
            url::emit_percent_encode(chunks, current, url::PercentOptions::path(), line)
        }
        "url.decode_form" => {
            url::emit_percent_decode(chunks, current, url::PercentOptions::form(), line)
        }
        "url.decode_rfc3986" => {
            url::emit_percent_decode(chunks, current, url::PercentOptions::rfc3986(), line)
        }
        "str_trim_chars" => strings::emit_trim_chars(
            chunks,
            current,
            argc,
            strings::TrimOptions::both(None),
            line,
        ),
        "str_trim_start_chars" => strings::emit_trim_chars(
            chunks,
            current,
            argc,
            strings::TrimOptions::start(None),
            line,
        ),
        "str_trim_end_chars" => {
            strings::emit_trim_chars(chunks, current, argc, strings::TrimOptions::end(None), line)
        }
        "str_scalar_last_index_of" => strings::emit_scalar_last_index_of(chunks, current, line),
        "str_code_point_at" => strings::emit_code_point_at(&mut chunks[current], line),
        "str_first_code_point" => strings::emit_first_code_point(&mut chunks[current], line),
        "str_to_upper" => strings::emit_to_upper(&mut chunks[current], line),
        "str_to_lower" => strings::emit_to_lower(&mut chunks[current], line),
        "str_trim" => strings::emit_trim(&mut chunks[current], line),
        "str_trim_start" => strings::emit_trim_start(&mut chunks[current], line),
        "str_trim_end" => strings::emit_trim_end(&mut chunks[current], line),
        "str_substring" => strings::emit_substring(&mut chunks[current], line),
        "str_split" => strings::emit_split(&mut chunks[current], line),
        "str_replace" => strings::emit_replace(&mut chunks[current], line),
        "str_repeat" => strings::emit_repeat(&mut chunks[current], line),
        "str_index_of" => strings::emit_index_of(&mut chunks[current], line),
        "str_last_index_of" => strings::emit_last_index_of(&mut chunks[current], line),
        "str_concat" => strings::emit_str_concat(&mut chunks[current], line),
        "str_from_char_code" => {
            let idx = chunks[current].add_import("ecma:string", "fromCharCode");
            chunks[current].emit_call(idx, argc as u8, line);
        }
        "str_char_code_at" => {
            let idx = chunks[current].add_import("wasm:js-string", "charCodeAt");
            chunks[current].emit_call(idx, 2, line);
        }
        "str_from_code_point" => strings::emit_from_code_point(&mut chunks[current], line),
        "str_char_at" => {
            let idx = chunks[current].add_import("ecma:string", "charAt");
            chunks[current].emit_call(idx, 2, line);
        }
        "str_starts_with" => {
            let idx = chunks[current].add_import("ecma:string", "startsWith");
            chunks[current].emit_call(idx, 2, line);
        }
        "str_ends_with" => {
            let idx = chunks[current].add_import("ecma:string", "endsWith");
            chunks[current].emit_call(idx, 2, line);
        }
        "str_contains" => {
            let idx = chunks[current].add_import("ecma:string", "includes");
            chunks[current].emit_call(idx, 2, line);
        }
        "str_includes" => {
            let idx = chunks[current].add_import("ecma:string", "indexOf");
            chunks[current].emit_call(idx, 2, line);
            crate::primitives::instructions::core_wasm::i32_const(&mut chunks[current], line, 0);
            ops::emit_dyn_ge(&mut chunks[current], line);
        }
        "str_compare" => {
            let idx = chunks[current].add_import("wasm:js-string", "compare");
            chunks[current].emit_call(idx, 2, line);
        }
        "str_pad_start" => {
            let idx = chunks[current].add_import("ecma:string", "padStart");
            chunks[current].emit_call(idx, 3, line);
        }
        "str_pad_end" => {
            let idx = chunks[current].add_import("ecma:string", "padEnd");
            chunks[current].emit_call(idx, 3, line);
        }

        // ── Dynamic ops (profile common:dyn_*) ──
        "dyn_eq" => ops::emit_dyn_eq(&mut chunks[current], line),
        "dyn_ne" => ops::emit_dyn_ne(&mut chunks[current], line),
        "dyn_to_bool" => ops::emit_dyn_to_bool(&mut chunks[current], line),

        // ── Ref ops (profile common:ref_*) ──
        "ref_is_array" => {
            let idx = chunks[current].add_import("ecma:array", "isArray");
            chunks[current].emit_call(idx, 1, line);
        }
        "ref_typeof" => {
            let idx = chunks[current].add_import("ecma:value", "typeof");
            chunks[current].emit_call(idx, 1, line);
        }
        "ref_eq" => {
            chunks[current].emit_op(Op::REF_EQ, line);
        }

        // ── Array ops (profile common:array_*) ──
        "array_push" => collections::emit_push(chunks, current, line),
        "array_pop" => collections::emit_pop(chunks, current, line),
        "array_shift" => collections::emit_shift(chunks, current, line),
        "array_length" => collections::emit_len(chunks, current, line),
        "array_reverse" => collections::emit_reverse(chunks, current, line),

        // ── Misc ──
        "to_int" => {
            let idx = chunks[current].add_import("ecma:number", "parseInt");
            chunks[current].emit_call(idx, 1, line);
        }
        "round" => {
            let idx = chunks[current].add_import("ecma:math", "round");
            chunks[current].emit_call(idx, 1, line);
        }
        "set_timer" => {
            let idx = chunks[current].add_import("web:timers", "setTimeout");
            chunks[current].emit_call(idx, 2, line);
        }

        _ => return false,
    }
    true
}

/// Handle common ops that need to register a host import in addition to
/// emitting bytecode. `import` is a callback that resolves an import to its
/// index (typically by adding to chunk[0]'s import table).
///
/// Returns `true` if `name` was recognized and emitted; on `true`, the
/// stack discipline matches the helper's contract (e.g. `threading.sleep`
/// leaves a `null` on the stack so the call site can drop it uniformly).
/// Returns `false` if the name is unknown OR doesn't need imports — call
/// `emit_common` for those.
pub fn emit_common_with_imports(
    name: &str,
    chunk: &mut Chunk,
    argc: u8,
    line: u32,
    mut import: impl FnMut(&str, &str) -> u16,
) -> bool {
    let _ = argc; // unused by current emits — kept for parity with `emit_common`
    match name {
        "threading.sleep" => {
            let wait_for_idx = import("wasi:clocks/monotonic-clock", "wait-for");
            thread_adapter::emit_thread_sleep(chunk, wait_for_idx, line);
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        _ => return false,
    }
    true
}
