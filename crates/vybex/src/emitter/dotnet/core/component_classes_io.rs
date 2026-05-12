use super::super::super::class_exports::DotnetClassExport;
use super::component_classes_common::static_only_class;
use vybe_bytecode::component_model::{ClassType, ConstructorDef, MethodBody, MethodDef};

pub(super) fn exports() -> Vec<DotnetClassExport> {
    vec![
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("StreamReader")
                .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.stream_reader_new"))
                .with_method(MethodDef::new("ReadLine", 0, MethodBody::Common("dotnet.stream_reader_read_line".into())))
                .with_method(MethodDef::new("ReadToEnd", 0, MethodBody::Common("dotnet.stream_reader_read_to_end".into())))
                .with_method(MethodDef::new("Close", 0, MethodBody::Common("dotnet.stream_reader_close".into())))
                .with_method(MethodDef::new("Dispose", 0, MethodBody::Common("dotnet.stream_reader_close".into()))),
        ),
        DotnetClassExport::new(
            "dotnet.System.IO",
            ClassType::new("StreamWriter")
                .with_constructor(ConstructorDef::new(1).with_common_backing("dotnet.stream_writer_new"))
                .with_method(MethodDef::new("Write", 1, MethodBody::Common("dotnet.stream_writer_write".into())))
                .with_method(MethodDef::new("WriteLine", 1, MethodBody::Common("dotnet.stream_writer_write_line".into())))
                .with_method(MethodDef::new("Flush", 0, MethodBody::Common("dotnet.stream_writer_flush".into())))
                .with_method(MethodDef::new("Close", 0, MethodBody::Common("dotnet.stream_writer_flush".into())))
                .with_method(MethodDef::new("Dispose", 0, MethodBody::Common("dotnet.stream_writer_flush".into()))),
        ),
        static_only_class(
            "dotnet.System.IO",
            "File",
            &[
                ("ReadAllText", 1, "wasi:filesystem", "readFile"),
                ("WriteAllText", 2, "wasi:filesystem", "writeFile"),
                ("AppendAllText", 2, "wasi:filesystem", "appendFile"),
                ("Exists", 1, "wasi:filesystem", "exists"),
                ("Delete", 1, "wasi:filesystem", "remove"),
                ("Copy", 2, "wasi:filesystem", "copy"),
                ("Move", 2, "wasi:filesystem", "rename"),
            ],
        ),
        static_only_class(
            "dotnet.System.IO",
            "Directory",
            &[
                ("CreateDirectory", 1, "wasi:filesystem", "mkdir"),
                ("GetFiles", 1, "wasi:filesystem", "listDir"),
            ],
        ),
        static_only_class(
            "dotnet.System.IO",
            "Path",
            &[
                ("Combine", 2, "wasi:filesystem", "pathCombine"),
                ("GetFileName", 1, "wasi:filesystem", "pathGetFileName"),
                ("GetExtension", 1, "wasi:filesystem", "pathGetExtension"),
                ("GetDirectoryName", 1, "wasi:filesystem", "pathGetDirectory"),
                ("GetFileNameWithoutExtension", 1, "wasi:filesystem", "pathGetFileNameWithoutExt"),
                ("ChangeExtension", 2, "wasi:filesystem", "pathChangeExtension"),
                ("GetFullPath", 1, "wasi:filesystem", "pathGetFullPath"),
                ("GetTempPath", 0, "wasi:filesystem", "pathGetTempPath"),
            ],
        ),
    ]
}
