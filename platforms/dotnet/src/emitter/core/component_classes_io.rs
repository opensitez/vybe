use super::super::super::class_exports::DotnetClassExport;
use vybe_runtime::component_model::{
    ClassType, ConstructorDef, HostTarget, MethodBody, MethodDef, PropertyDef,
};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        // `MemoryStream` — see `memory_stream_adapter`.
        // `Capacity`/`Length`/`Position`/`CanRead`/`CanWrite`/`CanSeek` are
        // COMPUTED accessors, declared in `tree_register::shared_emit_accessors`.
        DotnetClassExport::new(
            "dotnet.System.IO",
            {
                let mut ty = ClassType::new("MemoryStream");
                for argc in [0, 1, 2, 3, 4, 5] {
                    ty = ty.with_constructor(
                        ConstructorDef::new(argc).with_common_backing("dotnet.memory_stream_new"),
                    );
                }
                for (name, arity, emit) in [
                    ("ToArray", 0u8, "dotnet.ms_to_array"),
                    ("GetBuffer", 0, "dotnet.ms_get_buffer"),
                    ("TryGetBuffer", 0, "dotnet.ms_try_get_buffer"),
                    ("WriteByte", 1, "dotnet.ms_write_byte"),
                    ("ReadByte", 0, "dotnet.ms_read_byte"),
                    ("Write", 1, "dotnet.ms_write"),
                    ("Write", 3, "dotnet.ms_write"),
                    ("Read", 1, "dotnet.ms_read"),
                    ("Read", 3, "dotnet.ms_read"),
                    ("Seek", 2, "dotnet.ms_seek"),
                    ("SetLength", 1, "dotnet.ms_set_length"),
                    ("WriteTo", 1, "dotnet.ms_write_to"),
                    ("Close", 0, "dotnet.ms_close"),
                    ("Dispose", 0, "dotnet.ms_close"),
                    ("Flush", 0, "dotnet.ms_flush"),
                    // `CopyTo`/`CopyToAsync` — every arity .NET declares, so
                    // the buffer-size and cancellation-token overloads resolve
                    // rather than falling through to nothing.
                    ("CopyTo", 1, "dotnet.ms_copy_to"),
                    ("CopyTo", 2, "dotnet.ms_copy_to"),
                    ("CopyToAsync", 1, "dotnet.ms_copy_to_async"),
                    ("CopyToAsync", 2, "dotnet.ms_copy_to_async"),
                    ("CopyToAsync", 3, "dotnet.ms_copy_to_async"),
                    ("FlushAsync", 0, "dotnet.ms_flush_async"),
                    ("FlushAsync", 1, "dotnet.ms_flush_async"),
                    ("DisposeAsync", 0, "dotnet.ms_dispose_async"),
                    ("WriteAsync", 1, "dotnet.ms_write_async"),
                    ("WriteAsync", 2, "dotnet.ms_write_async"),
                    ("WriteAsync", 3, "dotnet.ms_write_async"),
                    ("WriteAsync", 4, "dotnet.ms_write_async"),
                    ("ReadAsync", 1, "dotnet.ms_read_async"),
                    ("ReadAsync", 2, "dotnet.ms_read_async"),
                    ("ReadAsync", 3, "dotnet.ms_read_async"),
                    ("ReadAsync", 4, "dotnet.ms_read_async"),
                    ("CanTimeout", 0, "dotnet.ms_can_timeout"),
                ] {
                    ty = ty.with_method(MethodDef::new(
                        name,
                        arity,
                        MethodBody::Common(emit.into()),
                    ));
                }
                ty
            },
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("FileStream")
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.file_open_read"),
                )
                .with_method(MethodDef::new(
                    "Read",
                    0,
                    MethodBody::Common("dotnet.stream_reader_read_to_end".into()),
                ))
                .with_method(MethodDef::new(
                    "Write",
                    1,
                    MethodBody::Common("dotnet.stream_writer_write".into()),
                ))
                .with_method(MethodDef::new(
                    "WriteByte",
                    1,
                    MethodBody::Common("dotnet.file_stream_write_byte".into()),
                ))
                .with_method(MethodDef::new(
                    "Close",
                    0,
                    MethodBody::Common("dotnet.stream_writer_flush".into()),
                ))
                .with_method(MethodDef::new(
                    "Dispose",
                    0,
                    MethodBody::Common("dotnet.stream_writer_flush".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            binary_reader_class(),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            binary_writer_class(),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("StreamReader")
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.stream_reader_new"),
                )
                .with_method(MethodDef::new(
                    "EndOfStream",
                    0,
                    MethodBody::Common("dotnet.stream_reader_at_end".into()),
                ))
                .with_method(MethodDef::new(
                    "ReadLine",
                    0,
                    MethodBody::Common("dotnet.stream_reader_read_line".into()),
                ))
                .with_method(MethodDef::new(
                    "ReadToEnd",
                    0,
                    MethodBody::Common("dotnet.stream_reader_read_to_end".into()),
                ))
                .with_method(MethodDef::new(
                    "Close",
                    0,
                    MethodBody::Common("dotnet.stream_reader_close".into()),
                ))
                .with_method(MethodDef::new(
                    "Dispose",
                    0,
                    MethodBody::Common("dotnet.stream_reader_close".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("StreamWriter")
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.stream_writer_new"),
                )
                .with_method(MethodDef::new(
                    "Write",
                    1,
                    MethodBody::Common("dotnet.stream_writer_write".into()),
                ))
                .with_method(MethodDef::new(
                    "WriteLine",
                    1,
                    MethodBody::Common("dotnet.stream_writer_write_line".into()),
                ))
                .with_method(MethodDef::new(
                    "Flush",
                    0,
                    MethodBody::Common("dotnet.stream_writer_flush".into()),
                ))
                .with_method(MethodDef::new(
                    "Close",
                    0,
                    MethodBody::Common("dotnet.stream_writer_flush".into()),
                ))
                .with_method(MethodDef::new(
                    "Dispose",
                    0,
                    MethodBody::Common("dotnet.stream_writer_flush".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("StringReader")
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.string_reader_new"),
                )
                .with_method(MethodDef::new(
                    "Peek",
                    0,
                    MethodBody::Common("dotnet.string_reader_peek".into()),
                ))
                .with_method(MethodDef::new(
                    "Read",
                    0,
                    MethodBody::Common("dotnet.string_reader_read".into()),
                ))
                .with_method(MethodDef::new(
                    "Read",
                    3,
                    MethodBody::Common("dotnet.string_reader_read_buffer".into()),
                ))
                .with_method(MethodDef::new(
                    "ReadBlock",
                    3,
                    MethodBody::Common("dotnet.string_reader_read_buffer".into()),
                ))
                .with_method(MethodDef::new(
                    "ReadLine",
                    0,
                    MethodBody::Common("dotnet.stream_reader_read_line".into()),
                ))
                .with_method(MethodDef::new(
                    "ReadLineAsync",
                    0,
                    MethodBody::Common("dotnet.stream_reader_read_line".into()),
                ))
                .with_method(MethodDef::new(
                    "ReadToEnd",
                    0,
                    MethodBody::Common("dotnet.stream_reader_read_to_end".into()),
                ))
                .with_method(MethodDef::new(
                    "Close",
                    0,
                    MethodBody::Common("dotnet.stream_reader_close".into()),
                ))
                .with_method(MethodDef::new(
                    "Dispose",
                    0,
                    MethodBody::Common("dotnet.stream_reader_close".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("StringWriter")
                .with_constructor(
                    ConstructorDef::new(0).with_common_backing("dotnet.string_writer_new"),
                )
                .with_constructor(
                    ConstructorDef::new(1).with_common_backing("dotnet.string_writer_new"),
                )
                .with_method(MethodDef::new(
                    "Write",
                    1,
                    MethodBody::Common("dotnet.stream_writer_write".into()),
                ))
                .with_method(MethodDef::new(
                    "Write",
                    3,
                    MethodBody::Common("dotnet.stream_writer_write_3".into()),
                ))
                .with_method(MethodDef::new(
                    "WriteLine",
                    1,
                    MethodBody::Common("dotnet.stream_writer_write_line".into()),
                ))
                .with_method(MethodDef::new(
                    "WriteLineAsync",
                    1,
                    MethodBody::Common("dotnet.stream_writer_write_line_async".into()),
                ))
                .with_method(MethodDef::new(
                    "GetStringBuilder",
                    0,
                    MethodBody::Common("dotnet.string_writer_get_string_builder".into()),
                ))
                .with_method(MethodDef::new(
                    "ToString",
                    0,
                    MethodBody::Common("dotnet.string_writer_to_string".into()),
                ))
                .with_method(MethodDef::new(
                    "Flush",
                    0,
                    MethodBody::Common("dotnet.string_writer_noop".into()),
                ))
                .with_method(MethodDef::new(
                    "Close",
                    0,
                    MethodBody::Common("dotnet.string_writer_noop".into()),
                ))
                .with_method(MethodDef::new(
                    "Dispose",
                    0,
                    MethodBody::Common("dotnet.string_writer_noop".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("File")
                // `dotnet.file_*`, not `filesystem.*` directly: the shared
                // lowering answers `null`/`false` — WASI's `result` in each
                // language's idiom — and .NET's idiom is an exception. Binding
                // straight to `filesystem.read_file` let `File.ReadAllText`
                // on a missing file return null, which prints as `0` and
                // concatenates as `""`. The WASI call is the same one; only the
                // failure conversion is .NET's.
                .with_method(MethodDef::static_method(
                    "ReadAllText",
                    1,
                    MethodBody::Common("dotnet.file_read_all_text".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteAllText",
                    2,
                    MethodBody::Common("dotnet.file_write_all_text".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "AppendAllText",
                    2,
                    MethodBody::Common("dotnet.file_append_all_text".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadAllBytes",
                    1,
                    MethodBody::Common("dotnet.file_read_all_bytes".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteAllBytes",
                    2,
                    MethodBody::Common("dotnet.file_write_all_bytes".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Exists",
                    1,
                    MethodBody::Common("filesystem.exists".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "Delete",
                    1,
                    MethodBody::Common("dotnet.directory_delete".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Delete",
                    2,
                    MethodBody::Common("dotnet.directory_delete".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Copy",
                    2,
                    MethodBody::Common("filesystem.copy".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "Move",
                    2,
                    MethodBody::Common("filesystem.rename".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "ReadAllLines",
                    1,
                    MethodBody::Common("dotnet.file_read_all_lines".into()),
                ))
                .with_method(MethodDef::static_method(
                    "WriteAllLines",
                    2,
                    MethodBody::Common("dotnet.file_write_all_lines".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Create",
                    1,
                    MethodBody::Common("dotnet.file_create".into()),
                ))
                .with_method(MethodDef::static_method(
                    "OpenRead",
                    1,
                    MethodBody::Common("dotnet.file_open_read".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("FileInfo").with_constructor(
                ConstructorDef::new(1).with_common_backing("dotnet.file_info_new"),
            ),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("Directory")
                .with_method(MethodDef::static_method(
                    "Exists",
                    1,
                    MethodBody::Common("filesystem.is_dir".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "CreateDirectory",
                    1,
                    MethodBody::Common("filesystem.mkdir".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "Delete",
                    1,
                    MethodBody::Common("filesystem.remove".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "Move",
                    2,
                    MethodBody::Common("filesystem.rename".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "GetFiles",
                    1,
                    MethodBody::Common("dotnet.directory_get_files".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetFiles",
                    2,
                    MethodBody::Common("dotnet.directory_get_files".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetDirectories",
                    1,
                    MethodBody::Common("dotnet.directory_get_directories".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetCurrentDirectory",
                    0,
                    MethodBody::Common("dotnet.get_current_directory".into()),
                )),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("Path")
                // ⚠The last `node:path` reference on the `Path` class, kept
                // only because `PropertyDef::getter` accepts a `HostTarget`
                // and nothing else — there is no `Common` getter, and widening
                // that enum in `vybe_runtime::component_model` to move one
                // constant character is not a trade worth making. Every METHOD
                // below is now on `primitives::paths`.
                .with_property(
                    PropertyDef::new("DirectorySeparatorChar")
                        .with_getter(HostTarget::new("node:path", "sep")),
                )
                .with_method(MethodDef::static_method(
                    "Combine",
                    2,
                    MethodBody::Common("dotnet.path_combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Combine",
                    3,
                    MethodBody::Common("dotnet.path_combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Combine",
                    4,
                    MethodBody::Common("dotnet.path_combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Join",
                    2,
                    MethodBody::Common("dotnet.path_combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Join",
                    3,
                    MethodBody::Common("dotnet.path_combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "Join",
                    4,
                    MethodBody::Common("dotnet.path_combine".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetFileName",
                    1,
                    MethodBody::Common("dotnet.path_get_file_name".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetExtension",
                    1,
                    MethodBody::Common("path.extension".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetDirectoryName",
                    1,
                    MethodBody::Common("dotnet.path_get_directory_name".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetFileNameWithoutExtension",
                    1,
                    MethodBody::Common("dotnet.path_get_file_name_without_extension".into()),
                ))
                .with_method(MethodDef::static_method(
                    "ChangeExtension",
                    2,
                    MethodBody::Common("dotnet.path_change_extension".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetFullPath",
                    1,
                    MethodBody::Common("dotnet.path_get_full_path".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetPathRoot",
                    1,
                    MethodBody::Common("dotnet.path_get_path_root".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetTempPath",
                    0,
                    MethodBody::Common("path.temp_path".to_string()),
                ))
                .with_method(MethodDef::static_method(
                    "GetTempFileName",
                    0,
                    MethodBody::Common("dotnet.path_get_temp_file_name".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetRandomFileName",
                    0,
                    MethodBody::Common("dotnet.path_get_random_file_name".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetInvalidFileNameChars",
                    0,
                    MethodBody::Common("dotnet.path_get_invalid_file_name_chars".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetInvalidPathChars",
                    0,
                    MethodBody::Common("dotnet.path_get_invalid_path_chars".into()),
                ))
                .with_method(MethodDef::static_method(
                    "HasExtension",
                    1,
                    MethodBody::Common("dotnet.path_has_extension".into()),
                ))
                .with_method(MethodDef::static_method(
                    "IsPathRooted",
                    1,
                    MethodBody::Common("dotnet.path_is_path_rooted".into()),
                ))
                .with_method(MethodDef::static_method(
                    "GetRelativePath",
                    2,
                    MethodBody::Common("dotnet.path_get_relative_path".into()),
                ))
                .with_method(MethodDef::static_method(
                    "TrimEndingDirectorySeparator",
                    1,
                    MethodBody::Common("dotnet.path_trim_ending_directory_separator".into()),
                )),
        ),
    ]
}

/// `BinaryWriter` — a cursor on a stream, not a file handle.
///
/// ⛔ Every `Write*` spelling below is arity 1 and they differ only by the
/// STATIC type of the argument, which a descriptor cannot express. `Write`
/// itself keeps .NET's default for an unannotated value; the VB walker rewrites
/// a call whose argument carries a width (a literal suffix, a `CShort(…)`) to
/// the matching spelling. See `binary_io_adapter`.
fn binary_writer_class() -> ClassType {
    let mut class = ClassType::new("BinaryWriter");
    // `(stream)`, `(stream, encoding)`, `(stream, encoding, leaveOpen)`.
    for arity in 1..=3u8 {
        class = class.with_constructor(
            ConstructorDef::new(arity).with_common_backing("dotnet.binary_writer_new"),
        );
    }
    for (name, emit) in [
        ("Write", "dotnet.binary_write_auto"),
        ("WriteBoolean", "dotnet.binary_write_bool"),
        ("WriteByte", "dotnet.binary_write_u8"),
        ("WriteSByte", "dotnet.binary_write_i8"),
        ("WriteInt16", "dotnet.binary_write_i16"),
        ("WriteUInt16", "dotnet.binary_write_u16"),
        ("WriteInt32", "dotnet.binary_write_i32"),
        ("WriteUInt32", "dotnet.binary_write_u32"),
        ("WriteInt64", "dotnet.binary_write_i64"),
        ("WriteUInt64", "dotnet.binary_write_u64"),
        ("WriteSingle", "dotnet.binary_write_f32"),
        ("WriteDouble", "dotnet.binary_write_f64"),
        ("WriteDecimal", "dotnet.binary_write_decimal"),
        ("WriteString", "dotnet.binary_write_string"),
        ("WriteChar", "dotnet.binary_write_char"),
        ("WriteBytes", "dotnet.binary_write_bytes"),
        ("Write7BitEncodedInt", "dotnet.binary_write_7bit"),
    ] {
        class = class.with_method(MethodDef::new(name, 1, MethodBody::Common(emit.into())));
    }
    class
        .with_method(MethodDef::new(
            "Seek",
            2,
            MethodBody::Common("dotnet.binary_seek".into()),
        ))
        .with_method(MethodDef::new(
            "Flush",
            0,
            MethodBody::Common("dotnet.binary_flush".into()),
        ))
        .with_method(MethodDef::new(
            "BaseStream",
            0,
            MethodBody::Common("dotnet.binary_base_stream".into()),
        ))
        .with_method(MethodDef::new(
            "Close",
            0,
            MethodBody::Common("dotnet.binary_close".into()),
        ))
        .with_method(MethodDef::new(
            "Dispose",
            0,
            MethodBody::Common("dotnet.binary_close".into()),
        ))
}

/// `BinaryReader` — every member is name-distinguished, so unlike the writer
/// nothing here needs the argument's type.
fn binary_reader_class() -> ClassType {
    let mut class = ClassType::new("BinaryReader");
    for arity in 1..=3u8 {
        class = class.with_constructor(
            ConstructorDef::new(arity).with_common_backing("dotnet.binary_reader_new"),
        );
    }
    for (name, emit) in [
        ("ReadBoolean", "dotnet.binary_read_bool"),
        ("ReadByte", "dotnet.binary_read_u8"),
        ("ReadSByte", "dotnet.binary_read_i8"),
        ("ReadInt16", "dotnet.binary_read_i16"),
        ("ReadUInt16", "dotnet.binary_read_u16"),
        ("ReadInt32", "dotnet.binary_read_i32"),
        ("ReadUInt32", "dotnet.binary_read_u32"),
        ("ReadInt64", "dotnet.binary_read_i64"),
        ("ReadUInt64", "dotnet.binary_read_u64"),
        ("ReadSingle", "dotnet.binary_read_f32"),
        ("ReadDouble", "dotnet.binary_read_f64"),
        ("ReadDecimal", "dotnet.binary_read_decimal"),
        ("ReadString", "dotnet.binary_read_string"),
        ("ReadChar", "dotnet.binary_read_char"),
        ("PeekChar", "dotnet.binary_peek_char"),
        ("Read7BitEncodedInt", "dotnet.binary_read_7bit"),
        ("BaseStream", "dotnet.binary_base_stream"),
        ("Close", "dotnet.binary_close"),
        ("Dispose", "dotnet.binary_close"),
    ] {
        class = class.with_method(MethodDef::new(name, 0, MethodBody::Common(emit.into())));
    }
    class
        .with_method(MethodDef::new(
            "ReadBytes",
            1,
            MethodBody::Common("dotnet.binary_read_bytes".into()),
        ))
        .with_method(MethodDef::new(
            "Read",
            1,
            MethodBody::Common("dotnet.binary_read_bytes".into()),
        ))
        // `Read(buffer, index, count)` fills the caller's array; `Read(count)`
        // above allocates one. Same name, different arity, different contract.
        .with_method(MethodDef::new(
            "Read",
            3,
            MethodBody::Common("dotnet.binary_read_into".into()),
        ))
}
