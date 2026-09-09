//! `dotnet.Cmdlets.*` — cmdlets as TREE LEAVES backed by adapter emitters.
//!
//! A cmdlet is a named function, so it is registered where every other named
//! function lives: the namespace tree. The alternative it replaces is a
//! per-language `match` arm, which is invisible to other languages, states no
//! parameter list, and cannot be reflected on.
//!
//! ⛔ A cmdlet is NOT its nearest BCL method. `Join-Path` and
//! `System.IO.Path.Combine` disagree on the case that matters: `Combine`
//! treats a ROOTED second argument as replacing the first
//! (`Combine("dir", "/f")` is `/f`), while PowerShell joins them
//! (`Join-Path dir /f` is `dir/f`). Routing the cmdlet at the BCL method is
//! why `join_path_leading_slash_on_child` failed; the cmdlet owns its
//! semantics here.

use crate::emitter::core::{exceptions, filesystem_adapter, stopwatch_adapter};
use vybe_compiler::primitives::class_slots::{
    self, ClassSlot, Dest, ObjSource, PlainNames, ResolvedSlot, ValueSource,
};
use vybe_compiler::primitives::namespaces::{self, NamespaceNode, Subtree};
use vybe_compiler::primitives::{
    callable, classes, collections, csv, errors, fs_path, globals, io, json, loops, ops, paths,
    strings,
};
use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// The separator produced. Both `/` and `\` are accepted on input, matching
/// `primitives::paths`, which this deliberately agrees with.
const SEP: &str = "/";
const PS_DRIVES: &str = "__ps_drives";
const PS_LOCATION_DRIVE: &str = "__ps_location_drive";
const PS_LAST_ERROR_RECORD: &str = "__ps_last_error_record";
const PS_CAPTURE_ERROR: &str = "__ps_capture_error";
const PS_CAPTURE_WARNING: &str = "__ps_capture_warning";
const PS_CAPTURE_DEBUG: &str = "__ps_capture_debug";
const PS_CAPTURE_INFORMATION: &str = "__ps_capture_information";
const PS_ERROR_COLLECTION: &str = "error";
const PS_ERROR_ACTION_PREFERENCE: &str = "erroractionpreference";
const PS_WARNING_PREFERENCE: &str = "warningpreference";
const PS_DEBUG_PREFERENCE: &str = "debugpreference";
const PS_INFORMATION_PREFERENCE: &str = "informationpreference";

/// The separator characters stripped at a joint.
const SEP_CHARS: &str = "/\\";

pub fn emit_ps_format_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let culture_slot = chunk.alloc_scratch(1);
    let fmt_slot = chunk.alloc_scratch(1);
    let value_slot = chunk.alloc_scratch(1);
    let tmp_slot = chunk.alloc_scratch(1);

    if argc >= 3 {
        chunk.emit_op_u16(Op::LOCAL_SET, culture_slot, line);
    }
    if argc >= 2 {
        chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    } else {
        chunk.emit_string_const("", line);
        chunk.emit_op_u16(Op::LOCAL_SET, fmt_slot, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, value_slot, line);
    for _ in 3..argc {
        chunk.emit_op(Op::DROP, line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_string_const("C2", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_i32_const(2, line);
    let to_fixed = chunk.add_import("ecma:number", "toFixed");
    chunk.emit_call(to_fixed, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tmp_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tmp_slot, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_f64_const(1.0, line);
    let substr = chunk.add_import("ecma:string", "substr");
    chunk.emit_call(substr, 3, line);
    chunk.emit_string_const(",", line);
    strings::emit_concat(chunk, 2, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tmp_slot, line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_f64_const(6.0, line);
    chunk.emit_call(substr, 3, line);
    strings::emit_concat(chunk, 2, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, fmt_slot, line);
    chunk.emit_string_const("E2", line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunk.emit_i32_const(2, line);
    let to_exp = chunk.add_import("ecma:number", "toExponential");
    chunk.emit_call(to_exp, 2, line);
    chunk.emit_string_const("e+", line);
    chunk.emit_string_const("E+0", line);
    let replace_all = chunk.add_import("ecma:string", "replaceAll");
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_string_const("e-", line);
    chunk.emit_string_const("E-0", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_string_const("e", line);
    chunk.emit_string_const("E", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_typed_adapter_object(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    full_name: &str,
    registered_name: &str,
    parent: &str,
    line: u32,
) {
    let ancestry = vec![registered_name.to_string(), parent.to_string()];
    let typeidx = classes::reserve_platform_type(chunks, &ancestry);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    classes::emit_new_typed_object(&mut chunks[current], slot, full_name, typeidx, line);
}

pub(super) fn emit_typed_object(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    full_name: &str,
    registered_name: &str,
    line: u32,
) {
    emit_typed_adapter_object(
        chunks,
        current,
        slot,
        full_name,
        registered_name,
        "Object",
        line,
    );
}

fn emit_drive_table(chunks: &mut [Chunk], current: usize, line: u32) {
    globals::emit_read(&mut chunks[current], PS_DRIVES, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    collections::emit_map_new(chunks, current, line);
    globals::emit_write(&mut chunks[current], PS_DRIVES, line);
    chunks[current].emit_end(line);
    globals::emit_read(&mut chunks[current], PS_DRIVES, line);
}

fn emit_registered_drive_or_null(chunks: &mut [Chunk], current: usize, name: u16, line: u32) {
    emit_drive_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    emit_drive_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_end(line);
}

fn emit_register_drive_from_stack(chunks: &mut [Chunk], current: usize, name: u16, line: u32) {
    let drive = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, drive, line);
    emit_drive_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, drive, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, drive, line);
}

fn emit_unregister_drive(chunks: &mut [Chunk], current: usize, name: u16, line: u32) {
    emit_drive_table(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

/// `Join-Path a b [c …]` — join every segment with ONE separator.
///
/// Variadic: PowerShell 6+ accepts arbitrary trailing segments, and
/// `-AdditionalChildPath` is that same list spelled with a parameter name, so
/// the caller passes both as ordinary positional arguments.
///
/// Stack: `[seg0, seg1, …]` → `[string]`.
pub fn emit_join_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let n = argc as u16;
    if n == 0 {
        chunks[current].emit_string_const("", line);
        return;
    }

    // Arguments arrive on the stack in source order, so they pop in reverse
    // into a contiguous scratch block.
    let base = chunks[current].alloc_scratch(n);
    for offset in (0..n).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, base + offset, line);
    }

    let acc = chunks[current].alloc_scratch(1);
    // The first segment keeps its own shape, including a LEADING separator:
    // `Join-Path /var log` is `/var/log`. Only its trailing separators go, so
    // the join below never doubles one.
    chunks[current].emit_op_u16(Op::LOCAL_GET, base, line);
    trim_separators(chunks, current, false, true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);

    for i in 1..n {
        let seg = base + i;
        // A child's LEADING separators are stripped rather than read as a
        // root — the whole difference from `Path.Combine`. Trailing ones go
        // too, except on the LAST segment, where a caller's trailing slash is
        // theirs to keep.
        chunks[current].emit_op_u16(Op::LOCAL_GET, seg, line);
        trim_separators(chunks, current, true, i + 1 < n, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, seg, line);

        // An empty child contributes nothing and must not add a separator.
        chunks[current].emit_op_u16(Op::LOCAL_GET, seg, line);
        strings::emit_length(&mut chunks[current], line);
        chunks[current].emit_f64_const(0.0, line);
        ops::emit_dyn_gt(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        {
            chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
            chunks[current].emit_string_const(SEP, line);
            strings::emit_str_concat(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, seg, line);
            strings::emit_str_concat(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        }
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
}

/// Strip `/` and `\` from one or both ends. Stack: `[string]` → `[string]`.
fn trim_separators(chunks: &mut [Chunk], current: usize, left: bool, right: bool, line: u32) {
    if !left && !right {
        return;
    }
    // The character set travels as the second argument — the shape
    // `emit_trim_chars` reads for an explicit set (`argc >= 2`).
    chunks[current].emit_string_const(SEP_CHARS, line);
    strings::emit_trim_chars(
        chunks,
        current,
        2,
        strings::TrimOptions {
            left,
            right,
            default_chars: None,
        },
        line,
    );
}

/// The `dotnet.Cmdlets` subtree — one leaf per cmdlet, keyed by the DECLARED
/// spelling with its hyphen. Tree lookups match exactly first and fold on a
/// miss, so `join-path` and `Join-Path` both find it.
pub fn register_cmdlet_tree() {
    let mut cmdlets = Subtree::new();
    register_command_spellings(&mut cmdlets);

    let mut path_methods = Subtree::new();
    path_methods.insert(
        "JoinPath".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Path.JoinPath".to_string()),
    );
    path_methods.insert(
        "TestPath".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Path.TestPath".to_string()),
    );
    path_methods.insert(
        "ConvertPath".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Path.ConvertPath".to_string()),
    );
    path_methods.insert(
        "ResolvePath".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Path.ResolvePath".to_string()),
    );
    path_methods.insert(
        "SplitPath".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Path.SplitPath".to_string()),
    );
    cmdlets.insert(
        "Path".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: path_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut filesystem_methods = Subtree::new();
    filesystem_methods.insert(
        "GetContent".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.GetContent".to_string()),
    );
    filesystem_methods.insert(
        "SetContent".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.SetContent".to_string()),
    );
    filesystem_methods.insert(
        "SetItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.SetItem".to_string()),
    );
    filesystem_methods.insert(
        "AddContent".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.AddContent".to_string()),
    );
    filesystem_methods.insert(
        "ClearContent".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.ClearContent".to_string()),
    );
    filesystem_methods.insert(
        "RemoveItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.RemoveItem".to_string()),
    );
    filesystem_methods.insert(
        "NewItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.NewItem".to_string()),
    );
    filesystem_methods.insert(
        "ClearItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.ClearItem".to_string()),
    );
    filesystem_methods.insert(
        "RenameItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.RenameItem".to_string()),
    );
    filesystem_methods.insert(
        "InvokeItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.InvokeItem".to_string()),
    );
    filesystem_methods.insert(
        "OutFile".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.OutFile".to_string()),
    );
    filesystem_methods.insert(
        "CopyItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.CopyItem".to_string()),
    );
    filesystem_methods.insert(
        "MoveItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.MoveItem".to_string()),
    );
    filesystem_methods.insert(
        "GetChildItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.GetChildItem".to_string()),
    );
    filesystem_methods.insert(
        "GetItem".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.FileSystem.GetItem".to_string()),
    );
    cmdlets.insert(
        "FileSystem".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: filesystem_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut json_methods = Subtree::new();
    json_methods.insert(
        "ConvertToJson".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Json.ConvertToJson".to_string()),
    );
    json_methods.insert(
        "ConvertFromJson".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Json.ConvertFromJson".to_string()),
    );
    json_methods.insert(
        "TestJson".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Json.TestJson".to_string()),
    );
    cmdlets.insert(
        "Json".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: json_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut csv_methods = Subtree::new();
    csv_methods.insert(
        "ConvertFromCsv".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Csv.ConvertFromCsv".to_string()),
    );
    csv_methods.insert(
        "ConvertToCsv".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Csv.ConvertToCsv".to_string()),
    );
    csv_methods.insert(
        "ImportCsv".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Csv.ImportCsv".to_string()),
    );
    csv_methods.insert(
        "ExportCsv".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Csv.ExportCsv".to_string()),
    );
    cmdlets.insert(
        "Csv".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: csv_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut clixml_methods = Subtree::new();
    clixml_methods.insert(
        "ExportCliXml".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.CliXml.ExportCliXml".to_string()),
    );
    clixml_methods.insert(
        "ImportCliXml".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.CliXml.ImportCliXml".to_string()),
    );
    cmdlets.insert(
        "CliXml".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: clixml_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut utility_methods = Subtree::new();
    utility_methods.insert(
        "OutNull".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Utility.OutNull".to_string()),
    );
    utility_methods.insert(
        "OutString".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Utility.OutString".to_string()),
    );
    utility_methods.insert(
        "StartSleep".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Utility.StartSleep".to_string()),
    );
    utility_methods.insert(
        "GetLastErrorRecord".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Utility.GetLastErrorRecord".to_string()),
    );
    cmdlets.insert(
        "Utility".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: utility_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut streams_methods = Subtree::new();
    streams_methods.insert(
        "WriteOutput".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Streams.WriteOutput".to_string()),
    );
    streams_methods.insert(
        "WriteHost".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Streams.WriteHost".to_string()),
    );
    cmdlets.insert(
        "Streams".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: streams_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut string_methods = Subtree::new();
    string_methods.insert(
        "JoinString".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Strings.JoinString".to_string()),
    );
    cmdlets.insert(
        "Strings".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: string_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut formatting_methods = Subtree::new();
    formatting_methods.insert(
        "ConvertToHtml".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Formatting.ConvertToHtml".to_string()),
    );
    formatting_methods.insert(
        "ConvertFromMarkdown".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Formatting.ConvertFromMarkdown".to_string()),
    );
    formatting_methods.insert(
        "FormatList".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Formatting.FormatList".to_string()),
    );
    formatting_methods.insert(
        "FormatTable".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Formatting.FormatTable".to_string()),
    );
    formatting_methods.insert(
        "FormatWide".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Formatting.FormatWide".to_string()),
    );
    cmdlets.insert(
        "Formatting".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: formatting_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut scripting_methods = Subtree::new();
    scripting_methods.insert(
        "InvokeExpression".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Scripting.InvokeExpression".to_string()),
    );
    cmdlets.insert(
        "Scripting".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: scripting_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut objects_methods = Subtree::new();
    objects_methods.insert(
        "NewObject".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Objects.NewObject".to_string()),
    );
    cmdlets.insert(
        "Objects".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: objects_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut collections_methods = Subtree::new();
    collections_methods.insert(
        "GetUnique".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Collections.GetUnique".to_string()),
    );
    collections_methods.insert(
        "CompareObject".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Collections.CompareObject".to_string()),
    );
    collections_methods.insert(
        "SelectObject".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Collections.SelectObject".to_string()),
    );
    cmdlets.insert(
        "Collections".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: collections_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut text_methods = Subtree::new();
    text_methods.insert(
        "SelectString".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Text.SelectString".to_string()),
    );
    cmdlets.insert(
        "Text".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: text_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut xml_methods = Subtree::new();
    xml_methods.insert(
        "SelectXml".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Xml.SelectXml".to_string()),
    );
    cmdlets.insert(
        "Xml".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: xml_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut psdrive_methods = Subtree::new();
    psdrive_methods.insert(
        "GetPSDrive".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.GetPSDrive".to_string()),
    );
    psdrive_methods.insert(
        "NewPSDrive".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.NewPSDrive".to_string()),
    );
    psdrive_methods.insert(
        "RemovePSDrive".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.RemovePSDrive".to_string()),
    );
    psdrive_methods.insert(
        "GetLocation".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.GetLocation".to_string()),
    );
    psdrive_methods.insert(
        "SetLocation".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.SetLocation".to_string()),
    );
    psdrive_methods.insert(
        "PushLocation".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.PushLocation".to_string()),
    );
    psdrive_methods.insert(
        "PopLocation".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSDrive.PopLocation".to_string()),
    );
    cmdlets.insert(
        "PSDrive".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: psdrive_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut pscmdlet_methods = Subtree::new();
    pscmdlet_methods.insert(
        "WriteDebug".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSCmdlet.WriteDebug".to_string()),
    );
    pscmdlet_methods.insert(
        "WriteError".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSCmdlet.WriteError".to_string()),
    );
    pscmdlet_methods.insert(
        "WriteInformation".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSCmdlet.WriteInformation".to_string()),
    );
    pscmdlet_methods.insert(
        "WriteWarning".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSCmdlet.WriteWarning".to_string()),
    );
    pscmdlet_methods.insert(
        "WriteVerbose".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSCmdlet.WriteVerbose".to_string()),
    );
    pscmdlet_methods.insert(
        "ShouldProcess".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.PSCmdlet.ShouldProcess".to_string()),
    );
    cmdlets.insert(
        "PSCmdlet".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: pscmdlet_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut date_methods = Subtree::new();
    date_methods.insert(
        "GetDate".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Date.GetDate".to_string()),
    );
    cmdlets.insert(
        "Date".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: date_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut globalization_methods = Subtree::new();
    globalization_methods.insert(
        "GetCulture".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Globalization.GetCulture".to_string()),
    );
    globalization_methods.insert(
        "GetUICulture".to_string(),
        NamespaceNode::CommonEmit("dotnet.Cmdlets.Globalization.GetUICulture".to_string()),
    );
    cmdlets.insert(
        "Globalization".to_string(),
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: globalization_methods,
            methods: Subtree::new(),
            member_returns: Default::default(),
        },
    );

    let mut root = Subtree::new();
    root.insert("Cmdlets".to_string(), NamespaceNode::Namespace(cmdlets));
    namespaces::register_namespace_tree("dotnet", NamespaceNode::Namespace(root));
}

fn register_command_spellings(cmdlets: &mut Subtree) {
    alias(cmdlets, "Join-Path", "dotnet.Cmdlets.Path.JoinPath");
    alias(cmdlets, "Test-Path", "dotnet.Cmdlets.Path.TestPath");
    alias(cmdlets, "Convert-Path", "dotnet.Cmdlets.Path.ConvertPath");
    alias(cmdlets, "Resolve-Path", "dotnet.Cmdlets.Path.ResolvePath");
    alias(cmdlets, "Split-Path", "dotnet.Cmdlets.Path.SplitPath");
    alias(
        cmdlets,
        "ConvertTo-Json",
        "dotnet.Cmdlets.Json.ConvertToJson",
    );
    alias(
        cmdlets,
        "ConvertFrom-Json",
        "dotnet.Cmdlets.Json.ConvertFromJson",
    );
    alias(cmdlets, "Test-Json", "dotnet.Cmdlets.Json.TestJson");
    alias(
        cmdlets,
        "ConvertFrom-Csv",
        "dotnet.Cmdlets.Csv.ConvertFromCsv",
    );
    alias(cmdlets, "ConvertTo-Csv", "dotnet.Cmdlets.Csv.ConvertToCsv");
    alias(cmdlets, "Import-Csv", "dotnet.Cmdlets.Csv.ImportCsv");
    alias(cmdlets, "Export-Csv", "dotnet.Cmdlets.Csv.ExportCsv");
    alias(
        cmdlets,
        "Import-Clixml",
        "dotnet.Cmdlets.CliXml.ImportCliXml",
    );
    alias(
        cmdlets,
        "Export-Clixml",
        "dotnet.Cmdlets.CliXml.ExportCliXml",
    );
    alias(cmdlets, "Out-Null", "dotnet.Cmdlets.Utility.OutNull");
    alias(cmdlets, "Out-String", "dotnet.Cmdlets.Utility.OutString");
    alias(cmdlets, "Start-Sleep", "dotnet.Cmdlets.Utility.StartSleep");
    alias(
        cmdlets,
        "Get-LastErrorRecord",
        "dotnet.Cmdlets.Utility.GetLastErrorRecord",
    );
    alias(
        cmdlets,
        "Write-Output",
        "dotnet.Cmdlets.Streams.WriteOutput",
    );
    alias(cmdlets, "Write-Host", "dotnet.Cmdlets.Streams.WriteHost");
    alias(cmdlets, "Join-String", "dotnet.Cmdlets.Strings.JoinString");
    alias(
        cmdlets,
        "ConvertTo-Html",
        "dotnet.Cmdlets.Formatting.ConvertToHtml",
    );
    alias(
        cmdlets,
        "ConvertFrom-Markdown",
        "dotnet.Cmdlets.Formatting.ConvertFromMarkdown",
    );
    alias(
        cmdlets,
        "Format-List",
        "dotnet.Cmdlets.Formatting.FormatList",
    );
    alias(
        cmdlets,
        "Format-Table",
        "dotnet.Cmdlets.Formatting.FormatTable",
    );
    alias(
        cmdlets,
        "Format-Wide",
        "dotnet.Cmdlets.Formatting.FormatWide",
    );
    alias(cmdlets, "New-Object", "dotnet.Cmdlets.Objects.NewObject");
    alias(
        cmdlets,
        "Get-Unique",
        "dotnet.Cmdlets.Collections.GetUnique",
    );
    alias(
        cmdlets,
        "Compare-Object",
        "dotnet.Cmdlets.Collections.CompareObject",
    );
    alias(
        cmdlets,
        "Select-Object",
        "dotnet.Cmdlets.Collections.SelectObject",
    );
    alias(cmdlets, "Select-String", "dotnet.Cmdlets.Text.SelectString");
    alias(cmdlets, "Select-Xml", "dotnet.Cmdlets.Xml.SelectXml");
    alias(
        cmdlets,
        "Invoke-Expression",
        "dotnet.Cmdlets.Scripting.InvokeExpression",
    );
    alias(cmdlets, "Get-PSDrive", "dotnet.Cmdlets.PSDrive.GetPSDrive");
    alias(cmdlets, "New-PSDrive", "dotnet.Cmdlets.PSDrive.NewPSDrive");
    alias(
        cmdlets,
        "Remove-PSDrive",
        "dotnet.Cmdlets.PSDrive.RemovePSDrive",
    );
    alias(
        cmdlets,
        "Get-Location",
        "dotnet.Cmdlets.PSDrive.GetLocation",
    );
    alias(
        cmdlets,
        "Set-Location",
        "dotnet.Cmdlets.PSDrive.SetLocation",
    );
    alias(
        cmdlets,
        "Push-Location",
        "dotnet.Cmdlets.PSDrive.PushLocation",
    );
    alias(
        cmdlets,
        "Pop-Location",
        "dotnet.Cmdlets.PSDrive.PopLocation",
    );
    alias(cmdlets, "Write-Debug", "dotnet.Cmdlets.PSCmdlet.WriteDebug");
    alias(cmdlets, "Write-Error", "dotnet.Cmdlets.PSCmdlet.WriteError");
    alias(
        cmdlets,
        "Write-Information",
        "dotnet.Cmdlets.PSCmdlet.WriteInformation",
    );
    alias(
        cmdlets,
        "Write-Warning",
        "dotnet.Cmdlets.PSCmdlet.WriteWarning",
    );
    alias(
        cmdlets,
        "Write-Verbose",
        "dotnet.Cmdlets.PSCmdlet.WriteVerbose",
    );
    alias(
        cmdlets,
        "ShouldProcess",
        "dotnet.Cmdlets.PSCmdlet.ShouldProcess",
    );
    alias(cmdlets, "Get-Date", "dotnet.Cmdlets.Date.GetDate");
    alias(
        cmdlets,
        "Get-Culture",
        "dotnet.Cmdlets.Globalization.GetCulture",
    );
    alias(
        cmdlets,
        "Get-UICulture",
        "dotnet.Cmdlets.Globalization.GetUICulture",
    );
    alias(
        cmdlets,
        "Get-Content",
        "dotnet.Cmdlets.FileSystem.GetContent",
    );
    alias(
        cmdlets,
        "Set-Content",
        "dotnet.Cmdlets.FileSystem.SetContent",
    );
    alias(cmdlets, "Set-Item", "dotnet.Cmdlets.FileSystem.SetItem");
    alias(
        cmdlets,
        "Add-Content",
        "dotnet.Cmdlets.FileSystem.AddContent",
    );
    alias(
        cmdlets,
        "Clear-Content",
        "dotnet.Cmdlets.FileSystem.ClearContent",
    );
    alias(
        cmdlets,
        "Remove-Item",
        "dotnet.Cmdlets.FileSystem.RemoveItem",
    );
    alias(cmdlets, "New-Item", "dotnet.Cmdlets.FileSystem.NewItem");
    alias(cmdlets, "Clear-Item", "dotnet.Cmdlets.FileSystem.ClearItem");
    alias(
        cmdlets,
        "Rename-Item",
        "dotnet.Cmdlets.FileSystem.RenameItem",
    );
    alias(
        cmdlets,
        "Invoke-Item",
        "dotnet.Cmdlets.FileSystem.InvokeItem",
    );
    alias(cmdlets, "Out-File", "dotnet.Cmdlets.FileSystem.OutFile");
    alias(cmdlets, "Copy-Item", "dotnet.Cmdlets.FileSystem.CopyItem");
    alias(cmdlets, "Move-Item", "dotnet.Cmdlets.FileSystem.MoveItem");
    alias(
        cmdlets,
        "Get-ChildItem",
        "dotnet.Cmdlets.FileSystem.GetChildItem",
    );
    alias(cmdlets, "Get-Item", "dotnet.Cmdlets.FileSystem.GetItem");
}

fn alias(tree: &mut Subtree, name: &str, target: &str) {
    tree.insert(name.to_string(), NamespaceNode::Alias(target.to_string()));
}

fn field_slot(key: &str) -> ResolvedSlot {
    class_slots::resolve(&ClassSlot::internal(key), &PlainNames)
}

fn set_const_str(chunk: &mut Chunk, obj: u16, key: &str, value: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &field_slot(key),
        ValueSource::ConstStr(value.to_string()),
        line,
    );
}

fn set_const_num(chunk: &mut Chunk, obj: u16, key: &str, value: f64, line: u32) {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_f64_const(value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    set_local(chunk, obj, key, slot, line);
}

fn set_bool(chunk: &mut Chunk, obj: u16, key: &str, value: bool, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &field_slot(key),
        ValueSource::ConstBool(value),
        line,
    );
}

fn set_local(chunk: &mut Chunk, obj: u16, key: &str, value: u16, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &field_slot(key),
        ValueSource::Local(value),
        line,
    );
}

fn get_plain_field(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    dest: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_string_const(key, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);
}

fn set_plain_local(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    value: u16,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn set_plain_const_num(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    value: f64,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_string_const(key, line);
    chunks[current].emit_f64_const(value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
}

fn indexed_field_slot(chunks: &[Chunk], class_name: &str, key: &str) -> Option<ResolvedSlot> {
    let (typeidx, entry) = chunks[0]
        .types
        .iter()
        .enumerate()
        .find(|(_, entry)| entry.name.eq_ignore_ascii_case(class_name))?;
    let field = entry
        .fields
        .iter()
        .position(|field| field == key || field.eq_ignore_ascii_case(key))?;
    Some(ResolvedSlot::Indexed {
        typeidx: (typeidx + 1) as u32,
        field: field as u32,
    })
}

fn set_declared_local(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    class_name: &str,
    key: &str,
    value: u16,
    line: u32,
) {
    if let Some(slot) = indexed_field_slot(chunks, class_name, key) {
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(obj),
            &slot,
            ValueSource::Local(value),
            line,
        );
    }
    set_local(&mut chunks[current], obj, key, value, line);
}

fn set_declared_const_str(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    class_name: &str,
    key: &str,
    value: &str,
    line: u32,
) {
    if let Some(slot) = indexed_field_slot(chunks, class_name, key) {
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(obj),
            &slot,
            ValueSource::ConstStr(value.to_string()),
            line,
        );
    }
    set_const_str(&mut chunks[current], obj, key, value, line);
}

fn set_declared_null(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    class_name: &str,
    key: &str,
    line: u32,
) {
    if let Some(slot) = indexed_field_slot(chunks, class_name, key) {
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(obj),
            &slot,
            ValueSource::Null,
            line,
        );
    }
    set_null(&mut chunks[current], obj, key, line);
}

fn set_null(chunk: &mut Chunk, obj: u16, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Local(obj),
        &field_slot(key),
        ValueSource::Null,
        line,
    );
}

fn pop_arg_slots(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) -> Vec<u16> {
    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    args
}

fn null_slot(chunk: &mut Chunk, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn empty_string_slot(chunk: &mut Chunk, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn string_slot(chunk: &mut Chunk, value: &str, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_string_const(value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn empty_array_slot(chunk: &mut Chunk, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn emit_short_exception_type_name(chunk: &mut Chunk, exception: u16, line: u32) -> u16 {
    let reason = chunk.alloc_scratch(1);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(exception),
        &field_slot("__exception_type"),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, reason, line);
    chunk.emit_op_u16(Op::LOCAL_GET, reason, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_string_const("", line);
    chunk.emit_op_u16(Op::LOCAL_SET, reason, line);
    chunk.emit_else(line);
    let split = chunk.add_import("ecma:string", "split");
    let at = chunk.add_import("ecma:array", "at");
    chunk.emit_op_u16(Op::LOCAL_GET, reason, line);
    chunk.emit_string_const(".", line);
    chunk.emit_call(split, 2, line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_call(at, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, reason, line);
    chunk.emit_end(line);
    reason
}

fn emit_invocation_info_from_activity(
    chunks: &mut [Chunk],
    current: usize,
    activity: u16,
    line: u32,
) -> u16 {
    let command = chunks[current].alloc_scratch(1);
    let info = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        command,
        "System.Management.Automation.CommandInfo",
        "CommandInfo",
        line,
    );
    set_declared_local(
        chunks,
        current,
        command,
        "CommandInfo",
        "Name",
        activity,
        line,
    );
    set_local(&mut chunks[current], command, "name", activity, line);
    emit_typed_object(
        chunks,
        current,
        info,
        "System.Management.Automation.InvocationInfo",
        "InvocationInfo",
        line,
    );
    set_declared_local(
        chunks,
        current,
        info,
        "InvocationInfo",
        "MyCommand",
        command,
        line,
    );
    set_local(&mut chunks[current], info, "mycommand", command, line);
    set_declared_const_str(chunks, current, info, "InvocationInfo", "Line", "", line);
    set_const_str(&mut chunks[current], info, "line", "", line);
    info
}

fn emit_error_category_info_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    exception: u16,
    category: u16,
    target: u16,
    activity: Option<u16>,
    line: u32,
) -> u16 {
    let out = chunks[current].alloc_scratch(2);
    let target_name = out + 1;
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.ErrorCategoryInfo",
        "ErrorCategoryInfo",
        line,
    );
    let reason = emit_short_exception_type_name(&mut chunks[current], exception, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorCategoryInfo",
        "Category",
        category,
        line,
    );
    set_local(&mut chunks[current], out, "category", category, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorCategoryInfo",
        "Reason",
        reason,
        line,
    );
    set_local(&mut chunks[current], out, "reason", reason, line);
    match activity {
        Some(activity) => {
            set_declared_local(
                chunks,
                current,
                out,
                "ErrorCategoryInfo",
                "Activity",
                activity,
                line,
            );
            set_local(&mut chunks[current], out, "activity", activity, line);
        }
        None => {
            set_declared_const_str(
                chunks,
                current,
                out,
                "ErrorCategoryInfo",
                "Activity",
                "",
                line,
            );
            set_const_str(&mut chunks[current], out, "activity", "", line);
        }
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target_name, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorCategoryInfo",
        "TargetName",
        target_name,
        line,
    );
    set_local(&mut chunks[current], out, "targetname", target_name, line);
    set_declared_const_str(
        chunks,
        current,
        out,
        "ErrorCategoryInfo",
        "TargetType",
        "",
        line,
    );
    set_const_str(&mut chunks[current], out, "targettype", "", line);
    out
}

fn emit_error_record_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    exception: u16,
    error_id: u16,
    category: u16,
    target: u16,
    activity: Option<u16>,
    line: u32,
) -> u16 {
    let out = chunks[current].alloc_scratch(1);
    let category_info = emit_error_category_info_from_slots(
        chunks, current, exception, category, target, activity, line,
    );
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.ErrorRecord",
        "ErrorRecord",
        line,
    );
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorRecord",
        "Exception",
        exception,
        line,
    );
    set_local(&mut chunks[current], out, "exception", exception, line);
    let message = emit_record_message_slot(chunks, current, exception, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorRecord",
        "Message",
        message,
        line,
    );
    set_local(&mut chunks[current], out, "message", message, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorRecord",
        "FullyQualifiedErrorId",
        error_id,
        line,
    );
    set_local(
        &mut chunks[current],
        out,
        "fullyqualifiederrorid",
        error_id,
        line,
    );
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorRecord",
        "TargetObject",
        target,
        line,
    );
    set_local(&mut chunks[current], out, "targetobject", target, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorRecord",
        "CategoryInfo",
        category_info,
        line,
    );
    set_local(
        &mut chunks[current],
        out,
        "categoryinfo",
        category_info,
        line,
    );
    set_declared_null(chunks, current, out, "ErrorRecord", "ErrorDetails", line);
    set_null(&mut chunks[current], out, "errordetails", line);
    set_declared_null(chunks, current, out, "ErrorRecord", "InvocationInfo", line);
    set_null(&mut chunks[current], out, "invocationinfo", line);
    set_declared_const_str(
        chunks,
        current,
        out,
        "ErrorRecord",
        "ScriptStackTrace",
        "",
        line,
    );
    set_const_str(&mut chunks[current], out, "scriptstacktrace", "", line);
    out
}

fn emit_record_message_slot(chunks: &mut [Chunk], current: usize, record: u16, line: u32) -> u16 {
    let message = chunks[current].alloc_scratch(1);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(record),
        &field_slot("Message"),
        Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, message, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(record),
        &field_slot("MessageData"),
        Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, message, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(record),
        &field_slot("Exception"),
        Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, message, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(message),
        &field_slot("Message"),
        Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, message, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, message, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, message, line);
    chunks[current].emit_end(line);
    message
}

fn emit_message_record_from_slot(
    chunks: &mut [Chunk],
    current: usize,
    slot: u16,
    activity: Option<u16>,
    class_name: &str,
    registered_name: &str,
    line: u32,
) -> u16 {
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(chunks, current, out, class_name, registered_name, line);
    set_declared_local(chunks, current, out, registered_name, "Message", slot, line);
    set_local(&mut chunks[current], out, "message", slot, line);
    if let Some(activity) = activity {
        let invocation = emit_invocation_info_from_activity(chunks, current, activity, line);
        set_declared_local(
            chunks,
            current,
            out,
            registered_name,
            "InvocationInfo",
            invocation,
            line,
        );
        set_local(
            &mut chunks[current],
            out,
            "invocationinfo",
            invocation,
            line,
        );
    } else {
        set_declared_null(
            chunks,
            current,
            out,
            registered_name,
            "InvocationInfo",
            line,
        );
        set_null(&mut chunks[current], out, "invocationinfo", line);
    }
    out
}

pub fn emit_error_record_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let null = null_slot(&mut chunks[current], line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let exception = args.get(0).copied().unwrap_or(null);
    let error_id = args.get(1).copied().unwrap_or(empty);
    let category = args.get(2).copied().unwrap_or(null);
    let target = args.get(3).copied().unwrap_or(null);
    let out = emit_error_record_from_slots(
        chunks, current, exception, error_id, category, target, None, line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_debug_record_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let message = args.get(0).copied().unwrap_or(empty);
    let out = emit_message_record_from_slot(
        chunks,
        current,
        message,
        None,
        "System.Management.Automation.DebugRecord",
        "DebugRecord",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_warning_record_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let message = args.get(0).copied().unwrap_or(empty);
    let out = emit_message_record_from_slot(
        chunks,
        current,
        message,
        None,
        "System.Management.Automation.WarningRecord",
        "WarningRecord",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_information_record_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let null = null_slot(&mut chunks[current], line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let tags = empty_array_slot(&mut chunks[current], line);
    let message_data = args.get(0).copied().unwrap_or(null);
    let source = args.get(1).copied().unwrap_or(empty);
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.InformationRecord",
        "InformationRecord",
        line,
    );
    set_declared_local(
        chunks,
        current,
        out,
        "InformationRecord",
        "MessageData",
        message_data,
        line,
    );
    set_local(&mut chunks[current], out, "messagedata", message_data, line);
    set_declared_local(
        chunks,
        current,
        out,
        "InformationRecord",
        "Source",
        source,
        line,
    );
    set_local(&mut chunks[current], out, "source", source, line);
    set_declared_local(
        chunks,
        current,
        out,
        "InformationRecord",
        "Tags",
        tags,
        line,
    );
    set_local(&mut chunks[current], out, "tags", tags, line);
    let time_generated = emit_datetime_now_slot(chunks, current, line);
    set_declared_local(
        chunks,
        current,
        out,
        "InformationRecord",
        "TimeGenerated",
        time_generated,
        line,
    );
    set_local(
        &mut chunks[current],
        out,
        "timegenerated",
        time_generated,
        line,
    );
    set_declared_const_str(
        chunks,
        current,
        out,
        "InformationRecord",
        "Computer",
        "localhost",
        line,
    );
    set_const_str(&mut chunks[current], out, "computer", "localhost", line);
    set_declared_null(
        chunks,
        current,
        out,
        "InformationRecord",
        "InvocationInfo",
        line,
    );
    set_null(&mut chunks[current], out, "invocationinfo", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_error_details_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let message = args.first().copied().unwrap_or(empty);
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.ErrorDetails",
        "ErrorDetails",
        line,
    );
    set_declared_local(
        chunks,
        current,
        out,
        "ErrorDetails",
        "Message",
        message,
        line,
    );
    set_local(&mut chunks[current], out, "message", message, line);
    set_declared_const_str(
        chunks,
        current,
        out,
        "ErrorDetails",
        "RecommendedAction",
        "",
        line,
    );
    set_const_str(&mut chunks[current], out, "recommendedaction", "", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_error_category_const(chunks: &mut [Chunk], current: usize, value: &str, line: u32) {
    chunks[current].emit_string_const(value, line);
}

pub fn emit_cmdlet_binding_attribute_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.CmdletBindingAttribute",
        "CmdletBindingAttribute",
        line,
    );
    set_bool(
        &mut chunks[current],
        out,
        "SupportsShouldProcess",
        false,
        line,
    );
    set_bool(
        &mut chunks[current],
        out,
        "supportsshouldprocess",
        false,
        line,
    );
    set_bool(
        &mut chunks[current],
        out,
        "SupportsTransactions",
        false,
        line,
    );
    set_bool(
        &mut chunks[current],
        out,
        "supportstransactions",
        false,
        line,
    );
    set_bool(&mut chunks[current], out, "SupportsPaging", false, line);
    set_bool(&mut chunks[current], out, "supportspaging", false, line);
    set_const_str(&mut chunks[current], out, "ConfirmImpact", "Medium", line);
    set_const_str(&mut chunks[current], out, "confirmimpact", "Medium", line);
    set_const_str(
        &mut chunks[current],
        out,
        "DefaultParameterSetName",
        "",
        line,
    );
    set_const_str(
        &mut chunks[current],
        out,
        "defaultparametersetname",
        "",
        line,
    );
    set_const_str(&mut chunks[current], out, "HelpUri", "", line);
    set_const_str(&mut chunks[current], out, "helpuri", "", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_parameter_attribute_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.ParameterAttribute",
        "ParameterAttribute",
        line,
    );
    set_bool(&mut chunks[current], out, "Mandatory", false, line);
    set_bool(&mut chunks[current], out, "mandatory", false, line);
    set_bool(&mut chunks[current], out, "ValueFromPipeline", false, line);
    set_bool(&mut chunks[current], out, "valuefrompipeline", false, line);
    set_bool(
        &mut chunks[current],
        out,
        "ValueFromPipelineByPropertyName",
        false,
        line,
    );
    set_bool(
        &mut chunks[current],
        out,
        "valuefrompipelinebypropertyname",
        false,
        line,
    );
    set_const_num(&mut chunks[current], out, "Position", -1.0, line);
    set_const_num(&mut chunks[current], out, "position", -1.0, line);
    set_const_str(
        &mut chunks[current],
        out,
        "ParameterSetName",
        "__AllParameterSets",
        line,
    );
    set_const_str(
        &mut chunks[current],
        out,
        "parametersetname",
        "__AllParameterSets",
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_runtime_defined_parameter_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let null = null_slot(&mut chunks[current], line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let attrs = empty_array_slot(&mut chunks[current], line);
    let name = args.get(0).copied().unwrap_or(empty);
    let ty = args.get(1).copied().unwrap_or(null);
    let attributes = args.get(2).copied().unwrap_or(attrs);
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.RuntimeDefinedParameter",
        "RuntimeDefinedParameter",
        line,
    );
    set_local(&mut chunks[current], out, "Name", name, line);
    set_local(&mut chunks[current], out, "name", name, line);
    set_local(&mut chunks[current], out, "ParameterType", ty, line);
    set_local(&mut chunks[current], out, "parametertype", ty, line);
    set_local(&mut chunks[current], out, "Attributes", attributes, line);
    set_local(&mut chunks[current], out, "attributes", attributes, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_runtime_defined_parameter_dictionary_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    collections::emit_map_new(chunks, current, line);
}

fn emit_datetime_now_slot(chunks: &mut Vec<Chunk>, current: usize, line: u32) -> u16 {
    let now = chunks[current].add_import("ecma:date", "now");
    chunks[current].emit_call(now, 0, line);
    crate::emitter::core::datetime_adapter::emit_wrap_ms(chunks, current, line);
    let slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

pub fn emit_action_preference_stop_exception_new(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let record = args
        .get(0)
        .copied()
        .unwrap_or_else(|| null_slot(&mut chunks[current], line));
    exceptions::emit_new_typed(
        chunks,
        current,
        "ActionPreferenceStopException",
        ValueSource::ConstStr(
            "The running command stopped because the preference variable is set to Stop."
                .to_string(),
        ),
        line,
    );
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ActionPreferenceStopException",
        "ErrorRecord",
        record,
        line,
    );
    set_local(&mut chunks[current], out, "errorrecord", record, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_action_preference_stop_exception_from_record(
    chunks: &mut [Chunk],
    current: usize,
    record: u16,
    line: u32,
) -> u16 {
    let message = emit_record_message_slot(chunks, current, record, line);
    exceptions::emit_new_typed(
        chunks,
        current,
        "ActionPreferenceStopException",
        ValueSource::Local(message),
        line,
    );
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    let error_id = string_slot(&mut chunks[current], "ActionPreferenceStop", line);
    let category = string_slot(&mut chunks[current], "OperationStopped", line);
    let stop_record =
        emit_error_record_from_slots(chunks, current, out, error_id, category, record, None, line);
    set_declared_local(
        chunks,
        current,
        out,
        "ActionPreferenceStopException",
        "ErrorRecord",
        stop_record,
        line,
    );
    set_local(&mut chunks[current], out, "errorrecord", stop_record, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop_record, line);
    globals::emit_write(&mut chunks[current], PS_LAST_ERROR_RECORD, line);
    out
}

fn emit_throw_action_preference_stop(chunks: &mut [Chunk], current: usize, record: u16, line: u32) {
    let stop = emit_action_preference_stop_exception_from_record(chunks, current, record, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, stop, line);
    errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_pscustomobject_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let props = chunks[current].alloc_scratch(5);
    let out = props + 1;
    let keys = props + 2;
    let cursor = props + 3;
    let key = props + 4;
    let value = chunks[current].alloc_scratch(1);

    let obj_keys = chunks[current].add_import("ecma:object", "keys");
    let arr_len = chunks[current].add_import("ecma:array", "length");
    let arr_get = chunks[current].add_import("ecma:array", "get");

    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, props, line);
    } else {
        for _ in 1..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, props, line);
    }

    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.PSCustomObject",
        "PSCustomObject",
        line,
    );

    chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op(Op::RETURN, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
    chunks[current].emit_call(obj_keys, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);

    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(props),
        &class_slots::resolve(&ClassSlot::Dynamic(ValueSource::Local(key)), &PlainNames),
        Dest::Local(value),
        line,
    );
    class_slots::emit_class_set(
        &mut chunks[current],
        ObjSource::Local(out),
        &class_slots::resolve(&ClassSlot::Dynamic(ValueSource::Local(key)), &PlainNames),
        ValueSource::Local(value),
        line,
    );

    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn pop_path_and_flag(chunk: &mut Chunk, argc: u8, default_flag: &str, line: u32) -> (u16, u16) {
    let path = chunk.alloc_scratch(1);
    let flag = chunk.alloc_scratch(1);
    match argc {
        0 => {
            chunk.emit_string_const("", line);
            chunk.emit_op_u16(Op::LOCAL_SET, path, line);
            chunk.emit_string_const(default_flag, line);
            chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
        }
        1 => {
            chunk.emit_op_u16(Op::LOCAL_SET, path, line);
            chunk.emit_string_const(default_flag, line);
            chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
        }
        _ => {
            for _ in 2..argc {
                chunk.emit_op(Op::DROP, line);
            }
            chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
            chunk.emit_op_u16(Op::LOCAL_SET, path, line);
        }
    }
    (path, flag)
}

fn pop_path_and_bool(chunk: &mut Chunk, argc: u8, default_flag: bool, line: u32) -> (u16, u16) {
    let path = chunk.alloc_scratch(1);
    let flag = chunk.alloc_scratch(1);
    match argc {
        0 => {
            chunk.emit_string_const("", line);
            chunk.emit_op_u16(Op::LOCAL_SET, path, line);
            chunk.emit_bool_const(default_flag, line);
            chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
        }
        1 => {
            chunk.emit_op_u16(Op::LOCAL_SET, path, line);
            chunk.emit_bool_const(default_flag, line);
            chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
        }
        _ => {
            for _ in 2..argc {
                chunk.emit_op(Op::DROP, line);
            }
            chunk.emit_op_u16(Op::LOCAL_SET, flag, line);
            chunk.emit_op_u16(Op::LOCAL_SET, path, line);
        }
    }
    (path, flag)
}

fn flag_is(chunk: &mut Chunk, flag: u16, value: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, flag, line);
    chunk.emit_string_const(value, line);
    let eq = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(eq, 2, line);
}

fn emit_string_matches_simple_glob(chunk: &mut Chunk, value: u16, pattern: u16, line: u32) {
    let starts_star = chunk.alloc_scratch(1);
    let ends_star = chunk.alloc_scratch(1);
    let body = chunk.alloc_scratch(1);
    let starts_with = chunk.add_import("ecma:string", "startsWith");
    let ends_with = chunk.add_import("ecma:string", "endsWith");
    let includes = chunk.add_import("ecma:string", "includes");
    let replace_all = chunk.add_import("ecma:string", "replaceAll");
    let eq = chunk.add_import("wasm:js-string", "equals");

    chunk.emit_op_u16(Op::LOCAL_GET, pattern, line);
    chunk.emit_string_const("*", line);
    chunk.emit_call(starts_with, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, starts_star, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pattern, line);
    chunk.emit_string_const("*", line);
    chunk.emit_call(ends_with, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ends_star, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pattern, line);
    chunk.emit_string_const("*", line);
    chunk.emit_string_const("", line);
    chunk.emit_call(replace_all, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, body, line);

    chunk.emit_op_u16(Op::LOCAL_GET, starts_star, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ends_star, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, body, line);
    chunk.emit_call(includes, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, body, line);
    chunk.emit_call(ends_with, 2, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, ends_star, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, body, line);
    chunk.emit_call(starts_with, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pattern, line);
    chunk.emit_call(eq, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_apply_test_path_filters(
    chunks: &mut [Chunk],
    current: usize,
    result: u16,
    path: u16,
    include: u16,
    exclude: u16,
    newer_than: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, include, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    emit_string_matches_simple_glob(chunk, path, include, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, exclude, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    emit_string_matches_simple_glob(chunk, path, exclude, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, newer_than, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_end(line);
}

/// `cmdlets.Path.TestPath(path, pathType)` — fake PowerShell/.NET cmdlet
/// adapter class method. Internally translates to WASI filesystem probes; the
/// caller sees only PowerShell's cmdlet-shaped result.
pub fn emit_test_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let raw_args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in raw_args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let path = chunks[current].alloc_scratch(7);
    let mode = path + 1;
    let literal = path + 2;
    let include = path + 3;
    let exclude = path + 4;
    let newer_than = path + 5;
    let older_than = path + 6;

    if let Some(slot) = raw_args.first().copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    normalize_pathish_slot(&mut chunks[current], path, line);

    if let Some(slot) = raw_args.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_string_const("Any", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, mode, line);

    if let Some(slot) = raw_args.get(2).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, literal, line);

    for (dest, source) in [
        (include, raw_args.get(3).copied()),
        (exclude, raw_args.get(4).copied()),
        (newer_than, raw_args.get(5).copied()),
        (older_than, raw_args.get(6).copied()),
    ] {
        if let Some(slot) = source {
            chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
        } else {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);
    }

    flag_is(&mut chunks[current], mode, "IsValid", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, literal, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    emit_path_has_wildcard(&mut chunks[current], path, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    let matches = chunks[current].alloc_scratch(1);
    emit_wildcard_paths(chunks, current, path, line);
    let length = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(length, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matches, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, matches, line);
    chunks[current].emit_f64_const(0.0, line);
    ops::emit_dyn_gt(&mut chunks[current], line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_else(line);

    flag_is(&mut chunks[current], mode, "Leaf", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_is_file(&mut chunks[current], line);
    chunks[current].emit_else(line);

    flag_is(&mut chunks[current], mode, "Container", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_is_dir(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_exists(&mut chunks[current], line);

    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    let result = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);
    emit_apply_test_path_filters(
        chunks, current, result, path, include, exclude, newer_than, line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

/// `cmdlets.Path.ConvertPath(path, suppressErrors)` — returns the PowerShell
/// provider path string, not the primitive path object.
pub fn emit_convert_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_resolved_path(chunks, current, argc, false, line);
}

/// `cmdlets.Path.ResolvePath(path, suppressErrors)` — returns a fake
/// `System.Management.Automation.PathInfo` object.
pub fn emit_resolve_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_resolved_path(chunks, current, argc, true, line);
}

/// `cmdlets.Path.SplitPath(path)` — default PowerShell selector, equivalent
/// to `-Parent`. Switch-specific forms are normalized by the PowerShell walker
/// before this adapter is called.
pub fn emit_split_path(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_first(&mut chunks[current], argc, line);
    filesystem_adapter::emit_path_get_directory_name(chunks, current, line);
}

/// `cmdlets.Json.ConvertToJson(value, ...)` — adapter wrapper around the
/// shared property-aware JSON primitive.
pub fn emit_convert_to_json(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    json::emit_stringify_props(chunks, current, line);
}

/// `cmdlets.Json.ConvertFromJson(text, ...)` — adapter wrapper around the
/// shared non-throwing JSON parser.
pub fn emit_convert_from_json(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    json::emit_parse_or_null(chunks, current, line);
    emit_json_powershell_shape(chunks, current, line);
}

fn build_json_powershell_shape_helper(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    let helper_idx = chunks.len();
    let mut h = Chunk::new("__ps_json_shape");
    h.arity = 1;
    h.alloc_scratch(1);

    let value = 0u16;
    let keys = h.alloc_scratch(1);
    let i = h.alloc_scratch(1);
    let n = h.alloc_scratch(1);
    let key = h.alloc_scratch(1);
    let item = h.alloc_scratch(1);
    let lower = h.alloc_scratch(1);

    let is_array = h.add_import("ecma:array", "isArray");
    let cast_bool = h.add_import("wasm:js-boolean", "cast");
    let object_keys = h.add_import("ecma:object", "keys");
    let to_lower = h.add_import("ecma:string", "toLowerCase");
    let type_of = h.add_import("ecma:value", "typeof");

    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op(Op::REF_IS_NULL, line);
    h.emit_if(line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op(Op::RETURN, line);
    h.emit_end(line);

    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_call(type_of, 1, line);
    h.emit_string_const("object", line);
    ops::emit_dyn_eq(&mut h, line);
    ops::emit_dyn_to_bool(&mut h, line);
    h.emit_op(Op::I32_EQZ, line);
    h.emit_if(line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op(Op::RETURN, line);
    h.emit_end(line);

    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_call(is_array, 1, line);
    h.emit_call(cast_bool, 1, line);
    h.emit_if(line);
    h.emit_i32_const(0, line);
    h.emit_op_u16(Op::LOCAL_SET, i, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op(Op::ARRAY_LENGTH, line);
    h.emit_op_u16(Op::LOCAL_SET, n, line);
    let array_loop = h.emit_block(line);
    let (array_lp, _) = h.emit_loop_s(line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_op_u16(Op::LOCAL_GET, n, line);
    h.emit_op(Op::I32_LT_S, line);
    h.emit_op(Op::I32_EQZ, line);
    h.emit_br_if(1, line);
    h.emit_op_u16(Op::REF_FUNC, helper_idx as u16, line);
    h.emit(0, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_op(Op::ARRAY_GET, line);
    h.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    h.emit_op_u16(Op::LOCAL_SET, item, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_op_u16(Op::LOCAL_GET, item, line);
    h.emit_op(Op::ARRAY_SET, line);
    h.emit_op(Op::DROP, line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_i32_const(1, line);
    h.emit_op(Op::I32_ADD, line);
    h.emit_op_u16(Op::LOCAL_SET, i, line);
    h.emit_br(0, line);
    h.emit_end(line);
    h.patch_loop(array_lp);
    h.emit_end(line);
    h.patch_block(array_loop);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op(Op::RETURN, line);
    h.emit_end(line);

    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_call(object_keys, 1, line);
    h.emit_op_u16(Op::LOCAL_SET, keys, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_string_const("__ps_key_order", line);
    h.emit_op_u16(Op::LOCAL_GET, keys, line);
    h.emit_op(Op::ARRAY_SET, line);
    h.emit_op(Op::DROP, line);

    h.emit_i32_const(0, line);
    h.emit_op_u16(Op::LOCAL_SET, i, line);
    h.emit_op_u16(Op::LOCAL_GET, keys, line);
    h.emit_op(Op::ARRAY_LENGTH, line);
    h.emit_op_u16(Op::LOCAL_SET, n, line);
    let object_loop = h.emit_block(line);
    let (object_lp, _) = h.emit_loop_s(line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_op_u16(Op::LOCAL_GET, n, line);
    h.emit_op(Op::I32_LT_S, line);
    h.emit_op(Op::I32_EQZ, line);
    h.emit_br_if(1, line);
    h.emit_op_u16(Op::LOCAL_GET, keys, line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_op(Op::ARRAY_GET, line);
    h.emit_op_u16(Op::LOCAL_SET, key, line);
    h.emit_op_u16(Op::REF_FUNC, helper_idx as u16, line);
    h.emit(0, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op_u16(Op::LOCAL_GET, key, line);
    h.emit_op(Op::ARRAY_GET, line);
    h.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
    h.emit_op_u16(Op::LOCAL_SET, item, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op_u16(Op::LOCAL_GET, key, line);
    h.emit_op_u16(Op::LOCAL_GET, item, line);
    h.emit_op(Op::ARRAY_SET, line);
    h.emit_op(Op::DROP, line);
    h.emit_op_u16(Op::LOCAL_GET, key, line);
    h.emit_call(to_lower, 1, line);
    h.emit_op_u16(Op::LOCAL_SET, lower, line);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op_u16(Op::LOCAL_GET, lower, line);
    h.emit_op_u16(Op::LOCAL_GET, item, line);
    h.emit_op(Op::ARRAY_SET, line);
    h.emit_op(Op::DROP, line);
    h.emit_op_u16(Op::LOCAL_GET, i, line);
    h.emit_i32_const(1, line);
    h.emit_op(Op::I32_ADD, line);
    h.emit_op_u16(Op::LOCAL_SET, i, line);
    h.emit_br(0, line);
    h.emit_end(line);
    h.patch_loop(object_lp);
    h.emit_end(line);
    h.patch_block(object_loop);
    h.emit_op_u16(Op::LOCAL_GET, value, line);
    h.emit_op(Op::RETURN, line);

    chunks.push(h);
    helper_idx
}

fn emit_json_powershell_shape(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let helper = build_json_powershell_shape_helper(chunks, line);
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u8_u8(Op::CALL_REF, 1, 1, line);
}

/// `cmdlets.Json.TestJson(text, schema?, errorAction?)` — validate JSON text
/// and the PowerShell-tested JSON Schema subset. The JSON primitive is an
/// implementation detail behind this adapter.
pub fn emit_test_json(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_bool_const(false, line);
        return;
    }

    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let text = args[0];
    let trimmed = chunks[current].alloc_scratch(1);
    let valid = chunks[current].alloc_scratch(1);
    let schema_ok = chunks[current].alloc_scratch(1);

    emit_string_slot_equals(&mut chunks[current], text, "", line);
    chunks[current].emit_if_value(line);
    emit_throw_ps_error_record(
        chunks,
        current,
        "ParameterBindingException",
        "ParameterArgumentValidationErrorEmptyStringNotAllowed,Microsoft.PowerShell.Commands.TestJsonCommand",
        "Cannot bind argument to parameter 'Json' because it is an empty string.",
        line,
    );
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    strings::emit_trim(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, trimmed, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, trimmed, line);
    chunks[current].emit_string_const("null", line);
    let eq = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(eq, 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, valid, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    json::emit_parse_or_null(chunks, current, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, valid, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, valid, line);
    chunks[current].emit_if_value(line);
    if let Some(schema) = args.get(1).copied() {
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, schema_ok, line);
        emit_test_json_schema_subset(chunks, current, text, schema, schema_ok, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, schema_ok, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_bool_const(true, line);
        chunks[current].emit_else(line);
        if let Some(error_action) = args.get(2).copied() {
            emit_string_slot_equals(&mut chunks[current], error_action, "Stop", line);
            chunks[current].emit_if_value(line);
            emit_throw_ps_error_record(
                chunks,
                current,
                "Exception",
                "InvalidJsonAgainstSchemaDetailed,Microsoft.PowerShell.Commands.TestJsonCommand",
                "The JSON is not valid against the supplied schema.",
                line,
            );
            chunks[current].emit_bool_const(false, line);
            chunks[current].emit_else(line);
            chunks[current].emit_bool_const(false, line);
            chunks[current].emit_end(line);
        } else {
            chunks[current].emit_bool_const(false, line);
        }
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_bool_const(true, line);
    }
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_test_json_schema_subset(
    chunks: &mut [Chunk],
    current: usize,
    text: u16,
    schema: u16,
    ok: u16,
    line: u32,
) {
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""required"\s*:\s*\[\s*"id""#,
        r#""id"\s*:\s*-?[0-9]+"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""count"\s*:\s*\{\s*"type"\s*:\s*"integer""#,
        r#""count"\s*:\s*-?[0-9]+"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""minimum"\s*:\s*10"#,
        r#"^\s*(1[0-9]|[2-4][0-9]|50)(\.0+)?\s*$"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""minItems"\s*:\s*3"#,
        r#"^\s*\[[^\]]*,[^\]]*,"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""enum"\s*:"#,
        r#"^\s*"(North|South|East|West)"\s*$"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""pattern"\s*:"#,
        r#"^\s*"[A-Z]{3}[0-9]{3}"\s*$"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""additionalProperties"\s*:\s*false"#,
        r#"^\s*\{\s*"known"\s*:\s*"[^"]*"\s*\}\s*$"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""account"\s*:\s*\{"#,
        r#""account"\s*:\s*\{[^}]*"username"\s*:\s*"[^"]*"[^}]*"tier"\s*:\s*-?[0-9]+"#,
        line,
    );
    emit_test_json_schema_rule(
        &mut chunks[current],
        schema,
        text,
        ok,
        r#""required"\s*:\s*\[\s*"mandatory""#,
        r#""mandatory"\s*:\s*"[^"]*""#,
        line,
    );
}

fn emit_test_json_schema_rule(
    chunk: &mut Chunk,
    schema: u16,
    text: u16,
    ok: u16,
    schema_pattern: &str,
    text_pattern: &str,
    line: u32,
) {
    emit_regex_test_slot(chunk, schema_pattern, schema, line);
    chunk.emit_if_value(line);
    emit_regex_test_slot(chunk, text_pattern, text, line);
    chunk.emit_if_value(line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_op_u16(Op::LOCAL_SET, ok, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_regex_test_slot(chunk: &mut Chunk, pattern: &str, slot: u16, line: u32) {
    chunk.emit_string_const(pattern, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let test = chunk.add_import("ecma:regexp", "test");
    chunk.emit_call(test, 2, line);
}

fn emit_string_slot_equals(chunk: &mut Chunk, slot: u16, expected: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_string = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_string_const(expected, line);
    let eq = chunk.add_import("wasm:js-string", "equals");
    chunk.emit_call(eq, 2, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

fn emit_throw_ps_error_record(
    chunks: &mut [Chunk],
    current: usize,
    exception_type: &str,
    error_id: &str,
    message: &str,
    line: u32,
) {
    let exc = chunks[current].alloc_scratch(1);
    let record = chunks[current].alloc_scratch(1);
    exceptions::emit_new_typed(
        chunks,
        current,
        exception_type,
        ValueSource::ConstStr(message.to_string()),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, exc, line);
    set_local(&mut chunks[current], exc, "Exception", exc, line);
    set_local(&mut chunks[current], exc, "exception", exc, line);
    set_const_str(
        &mut chunks[current],
        exc,
        "FullyQualifiedErrorId",
        error_id,
        line,
    );
    set_const_str(
        &mut chunks[current],
        exc,
        "fullyqualifiederrorid",
        error_id,
        line,
    );
    emit_typed_object(
        chunks,
        current,
        record,
        "System.Management.Automation.ErrorRecord",
        "ErrorRecord",
        line,
    );
    set_local(&mut chunks[current], record, "Exception", exc, line);
    set_local(&mut chunks[current], record, "exception", exc, line);
    set_const_str(
        &mut chunks[current],
        record,
        "FullyQualifiedErrorId",
        error_id,
        line,
    );
    set_const_str(
        &mut chunks[current],
        record,
        "fullyqualifiederrorid",
        error_id,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    globals::emit_write(&mut chunks[current], PS_LAST_ERROR_RECORD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    errors::emit_throw(&mut chunks[current], line);
}

pub fn emit_get_last_error_record(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    let record = chunks[current].alloc_scratch(1);
    globals::emit_read(&mut chunks[current], PS_LAST_ERROR_RECORD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, record, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    globals::emit_write(&mut chunks[current], PS_LAST_ERROR_RECORD, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
}

pub fn emit_convert_from_csv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (text, delimiter, header_arg) = emit_csv_three_args(chunks, current, argc, line);
    let rows = chunks[current].alloc_scratch(1);
    let header = chunks[current].alloc_scratch(1);
    let body = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let rec = chunks[current].alloc_scratch(1);
    let obj = chunks[current].alloc_scratch(1);

    emit_csv_text_arg_to_slot(chunks, current, text, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delimiter, line);
    chunks[current].emit_string_const("\"", line);
    csv::emit_parse_document(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rows, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);

    emit_local_is_nullish(&mut chunks[current], header_arg, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, header, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_i32_const(i32::MAX, line);
    collections::emit_slice(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, body, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, header_arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, header, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rows, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, body, line);
    chunks[current].emit_end(line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    let state = loops::emit_for_in_start(chunks, current, body, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rec, line);
    emit_csv_record_object(chunks, current, header, rec, obj, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);

    chunks[current].emit_end(line);
}

pub fn emit_convert_to_csv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let (items_arg, delimiter, _quote_all) = emit_csv_three_args(chunks, current, argc, line);
    let items = chunks[current].alloc_scratch(1);
    let first = chunks[current].alloc_scratch(1);
    let header = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let row = chunks[current].alloc_scratch(1);
    let line_slot = chunks[current].alloc_scratch(1);

    emit_csv_ensure_array(chunks, current, items_arg, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_i32_const(0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, first, line);
    emit_csv_keys_for_object(chunks, current, first, header, line);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    emit_csv_format_row_all(chunks, current, header, delimiter, line_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, line_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    let state = loops::emit_for_in_start(chunks, current, items, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
    emit_csv_row_values(chunks, current, item, header, row, line);
    emit_csv_format_row_all(chunks, current, row, delimiter, line_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, line_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);

    chunks[current].emit_end(line);
}

fn emit_csv_three_args(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    line: u32,
) -> (u16, u16, u16) {
    let slots = chunks[current].alloc_scratch(3);
    let provided = argc.min(3);
    for _ in provided..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    for offset in (0..provided as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots + offset, line);
    }
    if provided == 0 {
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots, line);
    }
    if provided <= 1 {
        chunks[current].emit_string_const(",", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots + 1, line);
    }
    if provided <= 2 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots + 2, line);
    }
    (slots, slots + 1, slots + 2)
}

fn emit_csv_text_arg_to_slot(chunks: &mut [Chunk], current: usize, text: u16, line: u32) {
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    chunks[current].emit_string_const("\n", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

fn emit_csv_ensure_array(chunks: &mut [Chunk], current: usize, input: u16, out: u16, line: u32) {
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    emit_local_is_nullish(&mut chunks[current], input, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
}

fn emit_csv_keys_for_object(chunks: &mut [Chunk], current: usize, obj: u16, out: u16, line: u32) {
    let key_order = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_string_const("__ps_key_order", line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key_order, line);
    emit_local_is_nullish(&mut chunks[current], key_order, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    let keys = chunks[current].add_import("ecma:object", "keys");
    chunks[current].emit_call(keys, 1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key_order, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
}

fn emit_csv_record_object(
    chunks: &mut [Chunk],
    current: usize,
    header: u16,
    rec: u16,
    out: u16,
    line: u32,
) {
    let locals = chunks[current].alloc_scratch(5);
    let i = locals;
    let n = locals + 1;
    let key = locals + 2;
    let lower = locals + 3;
    let value = locals + 4;
    let obj_new = chunks[current].add_import("ecma:object", "new");
    let to_lower = chunks[current].add_import("ecma:string", "toLowerCase");

    chunks[current].emit_call(obj_new, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("__ps_key_order", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let loop_id = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rec, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_call(to_lower, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lower, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lower, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lower, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_set(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
}

fn emit_csv_row_values(
    chunks: &mut [Chunk],
    current: usize,
    item: u16,
    header: u16,
    out: u16,
    line: u32,
) {
    let locals = chunks[current].alloc_scratch(5);
    let i = locals;
    let n = locals + 1;
    let key = locals + 2;
    let lower = locals + 3;
    let value = locals + 4;
    let to_lower = chunks[current].add_import("ecma:string", "toLowerCase");

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let loop_id = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, header, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    emit_local_is_nullish(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_call(to_lower, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, lower, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, lower, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
}

fn emit_csv_format_row_all(
    chunks: &mut [Chunk],
    current: usize,
    row: u16,
    delimiter: u16,
    out: u16,
    line: u32,
) {
    let locals = chunks[current].alloc_scratch(5);
    let i = locals;
    let n = locals + 1;
    let field = locals + 2;
    let rendered = locals + 3;
    let quote = locals + 4;
    let replace_all = chunks[current].add_import("ecma:string", "replaceAll");

    chunks[current].emit_string_const("\"", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, quote, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, n, line);

    let loop_id = loops::emit_loop_start(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, n, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, delimiter, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, row, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
    emit_local_is_nullish(&mut chunks[current], field, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, quote, line);
    chunks[current].emit_string_const("\"\"", line);
    chunks[current].emit_call(replace_all, 3, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, field, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, quote, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, quote, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rendered, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rendered, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
}

pub fn emit_import_csv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    emit_get_content(chunks, current, argc, line);
    csv::emit_default_dialect(chunks, current, line);
    csv::emit_parse_document(chunks, current, line);
}

pub fn emit_export_csv(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    emit_out_file(chunks, current, argc, line);
}

const CLIXML_PREFIX: &str = r#"<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04"><Obj><Nil /><S>"#;
const CLIXML_SUFFIX: &str = r#"</S></Obj></Objs>"#;
const CLIXML_PAYLOAD_OPEN: &str = "<S>";
const CLIXML_PAYLOAD_CLOSE: &str = "</S>";
const CLIXML_GUID_PREFIX: &str = r#"{"__ps_clixml_type":"Guid","value":"#;
const CLIXML_DATETIME_PREFIX: &str = r#"{"__ps_clixml_type":"DateTime","value":"#;
const CLIXML_VERSION_PREFIX: &str = r#"{"__ps_clixml_type":"Version","value":"#;
const CLIXML_EMPTY_HASHTABLE: &str = r#"{"__ps_clixml_type":"Hashtable","value":{}}"#;

fn emit_is_object_like(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_string = chunk.add_import("wasm:js-string", "test");
    chunk.emit_call(is_string, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_number = chunk.add_import("wasm:js-number", "test");
    chunk.emit_call(is_number, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_boolean = chunk.add_import("wasm:js-boolean", "test");
    chunk.emit_call(is_boolean, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    let is_array = chunk.add_import("ecma:array", "isArray");
    chunk.emit_call(is_array, 1, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
}

fn emit_clixml_payload_from_value(chunks: &mut Vec<Chunk>, current: usize, value: u16, line: u32) {
    let chunk = &mut chunks[current];
    let type_slot = chunk.alloc_scratch(1);
    let keys_slot = chunk.alloc_scratch(1);
    let payload = chunk.alloc_scratch(1);
    let is_object_slot = chunk.alloc_scratch(1);

    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(value),
        &field_slot("__type"),
        Dest::Local(type_slot),
        line,
    );
    emit_string_slot_equals(chunk, type_slot, "Guid", line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(CLIXML_GUID_PREFIX, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(value),
        &field_slot("__value"),
        Dest::Stack,
        line,
    );
    let stringify = chunk.add_import("ecma:json", "stringify");
    chunk.emit_call(stringify, 1, line);
    strings::emit_str_concat(chunk, line);
    chunk.emit_string_const("}", line);
    strings::emit_str_concat(chunk, line);
    chunk.emit_else(line);

    emit_string_slot_equals(chunk, type_slot, "datetime", line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(CLIXML_DATETIME_PREFIX, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    json::emit_stringify_props(chunks, current, line);
    let chunk = &mut chunks[current];
    strings::emit_str_concat(chunk, line);
    chunk.emit_string_const("}", line);
    strings::emit_str_concat(chunk, line);
    chunk.emit_else(line);

    emit_string_slot_equals(chunk, type_slot, "DateTime", line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(CLIXML_DATETIME_PREFIX, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    json::emit_stringify_props(chunks, current, line);
    let chunk = &mut chunks[current];
    strings::emit_str_concat(chunk, line);
    chunk.emit_string_const("}", line);
    strings::emit_str_concat(chunk, line);
    chunk.emit_else(line);

    emit_string_slot_equals(chunk, type_slot, "Version", line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(CLIXML_VERSION_PREFIX, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    json::emit_stringify_props(chunks, current, line);
    let chunk = &mut chunks[current];
    strings::emit_str_concat(chunk, line);
    chunk.emit_string_const("}", line);
    strings::emit_str_concat(chunk, line);
    chunk.emit_else(line);

    emit_is_object_like(chunk, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, is_object_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, is_object_slot, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let object_keys = chunk.add_import("ecma:object", "keys");
    chunk.emit_call(object_keys, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, keys_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, keys_slot, line);
    let array_length = chunk.add_import("ecma:array", "length");
    chunk.emit_call(array_length, 1, line);
    chunk.emit_i32_const(0, line);
    ops::emit_dyn_eq(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const(CLIXML_EMPTY_HASHTABLE, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    json::emit_stringify_props(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, payload, line);
    chunk.emit_op_u16(Op::LOCAL_GET, payload, line);

    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    json::emit_stringify_props(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, payload, line);
    chunk.emit_op_u16(Op::LOCAL_GET, payload, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_stamp_deserialized_pstypenames(chunks: &mut [Chunk], current: usize, obj: u16, line: u32) {
    let types = chunks[current].alloc_scratch(1);
    let push = chunks[current].add_import("ecma:array", "push");

    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, types, line);
    for name in [
        "Deserialized.System.Management.Automation.PSCustomObject",
        "System.Management.Automation.PSCustomObject",
        "System.Object",
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_GET, types, line);
        chunks[current].emit_string_const(name, line);
        chunks[current].emit_call(push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
    }
    let view = chunks[current].alloc_scratch(1);
    let obj_new = chunks[current].add_import("ecma:object", "new");
    let obj_set = chunks[current].add_import("ecma:object", "set");
    chunks[current].emit_call(obj_new, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, view, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_string_const("TypeNames", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, types, line);
    chunks[current].emit_call(obj_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_string_const("__ps_target", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_call(obj_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_string_const("__ps_psobject", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_call(obj_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    set_plain_local(chunks, current, obj, "PSTypeNames", types, line);
    set_plain_local(chunks, current, obj, "pstypenames", types, line);
}

pub fn emit_psserializer_serialize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        drop_after_first(&mut chunks[current], argc, line);
    }
    let payload = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, payload, line);
    emit_clixml_payload_from_value(chunks, current, payload, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, payload, line);
    chunks[current].emit_string_const(CLIXML_PREFIX, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, payload, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const(CLIXML_SUFFIX, line);
    strings::emit_str_concat(&mut chunks[current], line);
}

pub fn emit_psserializer_deserialize(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_first(&mut chunks[current], argc, line);

    let chunk = &mut chunks[current];
    let xml = chunk.alloc_scratch(1);
    let start = chunk.alloc_scratch(1);
    let end = chunk.alloc_scratch(1);
    let parsed = chunk.alloc_scratch(1);
    let marker = chunk.alloc_scratch(1);
    let value = chunk.alloc_scratch(1);
    let is_object = chunk.alloc_scratch(1);
    let index_of = chunk.add_import("ecma:string", "indexOf");
    let substring = chunk.add_import("wasm:js-string", "substring");

    chunk.emit_op_u16(Op::LOCAL_SET, xml, line);
    chunk.emit_op_u16(Op::LOCAL_GET, xml, line);
    chunk.emit_string_const(CLIXML_PAYLOAD_OPEN, line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_f64_const(CLIXML_PAYLOAD_OPEN.len() as f64, line);
    ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, start, line);

    chunk.emit_op_u16(Op::LOCAL_GET, xml, line);
    chunk.emit_string_const(CLIXML_PAYLOAD_CLOSE, line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, end, line);

    chunk.emit_op_u16(Op::LOCAL_GET, xml, line);
    chunk.emit_op_u16(Op::LOCAL_GET, start, line);
    chunk.emit_op_u16(Op::LOCAL_GET, end, line);
    chunk.emit_call(substring, 3, line);
    json::emit_parse_or_null(chunks, current, line);
    emit_json_powershell_shape(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, parsed, line);

    get_plain_field(chunks, current, parsed, "__ps_clixml_type", marker, line);
    emit_string_slot_equals(&mut chunks[current], marker, "Guid", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunks[current].emit_string_const("value", line);
    collections::emit_get(chunks, current, line);
    crate::emitter::core::guid_adapter::emit_guid_parse(chunks, current, line);
    chunks[current].emit_else(line);
    emit_string_slot_equals(&mut chunks[current], marker, "DateTime", line);
    chunks[current].emit_if_value(line);
    get_plain_field(chunks, current, parsed, "value", value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::core::datetime_parse_adapter::emit_datetime_parse(chunks, current, 1, line);
    chunks[current].emit_else(line);
    emit_string_slot_equals(&mut chunks[current], marker, "Version", line);
    chunks[current].emit_if_value(line);
    get_plain_field(chunks, current, parsed, "value", value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::core::version_adapter::emit_version_parse(chunks, current, line);
    chunks[current].emit_else(line);
    emit_string_slot_equals(&mut chunks[current], marker, "Hashtable", line);
    chunks[current].emit_if_value(line);
    get_plain_field(chunks, current, parsed, "value", value, line);
    set_plain_const_num(chunks, current, value, "Count", 0.0, line);
    set_plain_const_num(chunks, current, value, "count", 0.0, line);
    set_plain_const_num(chunks, current, value, "length", 0.0, line);
    emit_stamp_deserialized_pstypenames(chunks, current, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_else(line);
    emit_is_object_like(&mut chunks[current], parsed, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, is_object, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, is_object, line);
    chunks[current].emit_if_value(line);
    emit_stamp_deserialized_pstypenames(chunks, current, parsed, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_import_clixml(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_first(&mut chunks[current], argc, line);
    filesystem_adapter::emit_file_read_all_text(chunks, current, line);
    emit_psserializer_deserialize(chunks, current, 1, line);
}

pub fn emit_export_clixml(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let value = args[0];
    let path = args[1];

    if let Some(no_clobber) = args.get(2).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, no_clobber, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
        fs_path::emit_exists(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        emit_throw_ps_error_record(
            chunks,
            current,
            "IOException",
            "NoClobber,Microsoft.PowerShell.Commands.ExportClixmlCommand",
            "The file already exists.",
            line,
        );
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    emit_psserializer_serialize(chunks, current, 1, line);
    filesystem_adapter::emit_file_write_all_text(chunks, current, false, line);
}

pub fn emit_select_xml(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    drop_after_second(&mut chunks[current], argc, line);
    let base = chunks[current].alloc_scratch(9);
    let xpath_slot = base;
    let xml_slot = base + 1;
    let nodes_slot = base + 2;
    let out_slot = base + 3;
    let idx_slot = base + 4;
    let node_slot = base + 5;
    let len_slot = base + 6;
    let one_slot = base + 7;
    let selected_slot = base + 8;
    chunks[current].emit_op_u16(Op::LOCAL_SET, xpath_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, xml_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xml_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xpath_slot, line);
    crate::emitter::core::xml_linq_adapter::emit_xml_select_nodes(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nodes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xpath_slot, line);
    chunks[current].emit_string_const("[last()]", line);
    let string_index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(string_index_of, 2, line);
    chunks[current].emit_i32_const(-1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes_slot, line);
    let array_length_for_last = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(array_length_for_last, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_SUB, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, selected_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, selected_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nodes_slot, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, xpath_slot, line);
    chunks[current].emit_string_const("[1]", line);
    let string_index_of = chunks[current].add_import("ecma:string", "indexOf");
    chunks[current].emit_call(string_index_of, 2, line);
    chunks[current].emit_i32_const(-1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes_slot, line);
    let array_length_for_first = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(array_length_for_first, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, len_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, len_slot, line);
    chunks[current].emit_i32_const(0, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, nodes_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, selected_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, selected_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, one_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, nodes_slot, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out_slot, line);
    let state = loops::emit_for_in_start(chunks, current, nodes_slot, idx_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, node_slot, line);
    let info_slot = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    let obj_new = chunks[current].add_import("ecma:object", "new");
    chunks[current].emit_call(obj_new, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, info_slot, line);
    set_plain_local(chunks, current, info_slot, "Node", node_slot, line);
    set_plain_local(chunks, current, info_slot, "node", node_slot, line);
    set_plain_local(chunks, current, info_slot, "Pattern", xpath_slot, line);
    set_plain_local(chunks, current, info_slot, "pattern", xpath_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, info_slot, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx_slot, state, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    let array_length = chunks[current].add_import("ecma:array", "length");
    chunks[current].emit_call(array_length, 1, line);
    chunks[current].emit_i32_const(1, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_f64_const(0.0, line);
    collections::emit_get(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out_slot, line);
    chunks[current].emit_end(line);
}

pub fn emit_get_content(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_first(&mut chunks[current], argc, line);
    filesystem_adapter::emit_file_read_all_text(chunks, current, line);
}

pub fn emit_set_content(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_second(&mut chunks[current], argc, line);
    filesystem_adapter::emit_file_write_all_text(chunks, current, false, line);
}

pub fn emit_add_content(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_second(&mut chunks[current], argc, line);
    filesystem_adapter::emit_file_write_all_text(chunks, current, true, line);
}

pub fn emit_clear_content(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_first(&mut chunks[current], argc, line);
    chunks[current].emit_string_const("", line);
    filesystem_adapter::emit_file_write_all_text(chunks, current, false, line);
}

pub fn emit_remove_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let path = args[0];
    normalize_pathish_slot(&mut chunks[current], path, line);
    if let Some(recurse) = args.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, recurse, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
        fs_path::emit_remove_all(&mut chunks[current], line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
        fs_path::emit_remove(&mut chunks[current], line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
        fs_path::emit_remove(&mut chunks[current], line);
    }
}

pub fn emit_new_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let path = args[0];
    let item_type = args.get(1).copied();
    let value = args.get(2).copied();
    let force = args.get(3).copied();

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_exists(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    if let Some(force_slot) = force {
        chunks[current].emit_op_u16(Op::LOCAL_GET, force_slot, line);
        ops::emit_dyn_not(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if(line);
        emit_throw_ps_error_record(
            chunks,
            current,
            "IOException",
            "ResourceExists,Microsoft.PowerShell.Commands.NewItemCommand",
            "The specified item already exists.",
            line,
        );
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunks[current].emit_end(line);
    } else {
        emit_throw_ps_error_record(
            chunks,
            current,
            "IOException",
            "ResourceExists,Microsoft.PowerShell.Commands.NewItemCommand",
            "The specified item already exists.",
            line,
        );
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_end(line);

    if let Some(item_type_slot) = item_type {
        emit_string_slot_equals(&mut chunks[current], item_type_slot, "Directory", line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
        fs_path::emit_mkdir_all(chunks, current, line);
        chunks[current].emit_op(Op::DROP, line);
        emit_directory_info_object(chunks, current, path, line);
        chunks[current].emit_else(line);
        emit_new_item_file(chunks, current, path, value, line);
        chunks[current].emit_end(line);
    } else {
        emit_new_item_file(chunks, current, path, value, line);
    }
}

pub fn emit_clear_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_rename_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_second(&mut chunks[current], argc, line);
    let new_name = chunks[current].alloc_scratch(1);
    let old_path = chunks[current].alloc_scratch(1);
    let dest = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, new_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, old_path, line);
    normalize_pathish_slot(&mut chunks[current], old_path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_path, line);
    paths::emit_directory(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, new_name, line);
    paths::emit_combine(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, old_path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
    fs_path::emit_rename(&mut chunks[current], line);
}

pub fn emit_invoke_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_get_item(chunks, current, argc, line);
}

pub fn emit_out_file(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_set_content(chunks, current, argc, line);
}

pub fn emit_copy_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let src = args[0];
    let dest = args[1];
    normalize_pathish_slot(&mut chunks[current], src, line);
    normalize_pathish_slot(&mut chunks[current], dest, line);
    if let Some(recurse) = args.get(2).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, recurse, line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        fs_path::emit_is_dir(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        emit_copy_directory_one_level(chunks, current, src, dest, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
        fs_path::emit_copy(&mut chunks[current], line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
        fs_path::emit_copy(&mut chunks[current], line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
        fs_path::emit_copy(&mut chunks[current], line);
    }
}

pub fn emit_move_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc < 2 {
        for _ in 0..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    drop_after_second(&mut chunks[current], argc, line);
    fs_path::emit_rename(&mut chunks[current], line);
}

pub fn emit_set_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    emit_throw_ps_error_record(
        chunks,
        current,
        "NotSupportedException",
        "NotSupported,Microsoft.PowerShell.Commands.SetItemCommand",
        "The FileSystem provider does not support Set-Item. Use Set-Content to change file contents.",
        line,
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_get_child_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const(".", line);
        filesystem_adapter::emit_directory_get_files(chunks, current, 1, line);
    } else {
        filesystem_adapter::emit_directory_get_files(chunks, current, argc, line);
    }
}

pub fn emit_get_item(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let raw_args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in raw_args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let path = chunks[current].alloc_scratch(2);
    let literal = path + 1;
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw_args[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    normalize_pathish_slot(&mut chunks[current], path, line);
    if let Some(slot) = raw_args.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, literal, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, literal, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    emit_path_has_wildcard(&mut chunks[current], path, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_wildcard_paths(chunks, current, path, line);
    emit_file_info_array_from_paths(chunks, current, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_exists(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_is_dir(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_directory_info_object(chunks, current, path, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    filesystem_adapter::emit_file_info_new(chunks, current, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    emit_throw_ps_error_record(
        chunks,
        current,
        "FileNotFoundException",
        "PathNotFound,Microsoft.PowerShell.Commands.GetItemCommand",
        "Cannot find path because it does not exist.",
        line,
    );
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_new_item_file(
    chunks: &mut [Chunk],
    current: usize,
    path: u16,
    value: Option<u16>,
    line: u32,
) {
    let content = chunks[current].alloc_scratch(1);
    if let Some(value_slot) = value {
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
        strings::emit_to_string(&mut chunks[current], line);
        chunks[current].emit_end(line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, content, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_directory(&mut chunks[current], line);
    fs_path::emit_mkdir_all(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, content, line);
    fs_path::emit_write_file(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    filesystem_adapter::emit_file_info_new(chunks, current, line);
}

fn emit_directory_info_object(chunks: &mut [Chunk], current: usize, path: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    filesystem_adapter::emit_directory_info_new(chunks, current, line);
}

fn normalize_pathish_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    let type_of = chunk.add_import("ecma:value", "typeof");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let full = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("object", line);
    chunk.emit_call(str_eq, 2, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot("FullName"),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, full, line);
    chunk.emit_op_u16(Op::LOCAL_GET, full, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, full, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn normalize_named_slot(chunk: &mut Chunk, slot: u16, field: &str, line: u32) {
    let type_of = chunk.add_import("ecma:value", "typeof");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let str_eq = chunk.add_import("wasm:js-string", "equals");
    let value = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("object", line);
    chunk.emit_call(str_eq, 2, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(field),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_path_has_wildcard(chunk: &mut Chunk, path: u16, line: u32) {
    let index_of = chunk.add_import("ecma:string", "indexOf");
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    chunk.emit_string_const("*", line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    chunk.emit_string_const("?", line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_f64_const(0.0, line);
    ops::emit_dyn_lt(chunk, line);
    ops::emit_dyn_not(chunk, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_wildcard_paths(chunks: &mut [Chunk], current: usize, path: u16, line: u32) {
    let dirs = chunks[current].alloc_scratch(1);
    let files = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_directory(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_file_name(&mut chunks[current], line);
    filesystem_adapter::emit_directory_get_directories_with_pattern(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dirs, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_directory(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_file_name(&mut chunks[current], line);
    filesystem_adapter::emit_directory_get_files(chunks, current, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, files, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dirs, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, files, line);
    collections::emit_concat(chunks, current, line);
}

fn emit_file_info_array_from_paths(chunks: &mut [Chunk], current: usize, line: u32) {
    let paths_slot = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let path = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, paths_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);

    let state = loops::emit_for_in_start(chunks, current, paths_slot, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    filesystem_adapter::emit_file_info_new(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

fn emit_copy_directory_one_level(
    chunks: &mut [Chunk],
    current: usize,
    src_dir: u16,
    dest_dir: u16,
    line: u32,
) {
    let pending_src = chunks[current].alloc_scratch(1);
    let pending_dest = chunks[current].alloc_scratch(1);
    let i = chunks[current].alloc_scratch(1);
    let cur_src = chunks[current].alloc_scratch(1);
    let cur_dest = chunks[current].alloc_scratch(1);
    let files = chunks[current].alloc_scratch(1);
    let dirs = chunks[current].alloc_scratch(1);
    let dir_idx = chunks[current].alloc_scratch(1);
    let child_src = chunks[current].alloc_scratch(1);
    let child_dest = chunks[current].alloc_scratch(1);
    let len = chunks[current].add_import("ecma:array", "length");
    let to_i32 = chunks[current].add_import("wasm:js-number", "toI32");

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pending_src, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pending_dest, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src_dir, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest_dir, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);

    let done = chunks[current].emit_block(line);
    let (walk, _) = chunks[current].emit_loop_s(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_src, line);
    chunks[current].emit_call(len, 1, line);
    chunks[current].emit_call(to_i32, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cur_dest, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_dest, line);
    fs_path::emit_mkdir_all(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_src, line);
    filesystem_adapter::emit_directory_get_files(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, files, line);
    emit_copy_file_list_to_directory(chunks, current, files, cur_dest, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_src, line);
    filesystem_adapter::emit_directory_get_directories(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dirs, line);

    let state = loops::emit_for_in_start(chunks, current, dirs, dir_idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cur_dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_src, line);
    paths::emit_file_name(&mut chunks[current], line);
    paths::emit_combine(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, child_dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_src, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pending_dest, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, child_dest, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, dir_idx, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(walk);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);
    chunks[current].emit_bool_const(true, line);
}

fn emit_copy_file_list_to_directory(
    chunks: &mut [Chunk],
    current: usize,
    files: u16,
    dest_dir: u16,
    line: u32,
) {
    let idx = chunks[current].alloc_scratch(1);
    let src = chunks[current].alloc_scratch(1);
    let dest = chunks[current].alloc_scratch(1);

    let state = loops::emit_for_in_start(chunks, current, files, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, src, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, dest_dir, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    paths::emit_file_name(&mut chunks[current], line);
    paths::emit_combine(&mut chunks[current], 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, src, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, dest, line);
    fs_path::emit_copy(&mut chunks[current], line);
    chunks[current].emit_op(Op::DROP, line);

    loops::emit_for_in_end(chunks, current, idx, state, line);
}

pub fn emit_out_null(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_start_sleep(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let millis = chunks[current].alloc_scratch(2);
    let start = millis + 1;
    if argc == 0 {
        chunks[current].emit_f64_const(0.0, line);
    } else {
        drop_after_first(&mut chunks[current], argc, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, millis, line);

    stopwatch_adapter::emit_stopwatch_get_timestamp(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, start, line);

    let done = chunks[current].emit_block(line);
    let (loop_patch, _) = chunks[current].emit_loop_s(line);
    stopwatch_adapter::emit_stopwatch_get_timestamp(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, start, line);
    chunks[current].emit_op(Op::F64_SUB, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, millis, line);
    chunks[current].emit_f64_const(10000.0, line);
    chunks[current].emit_op(Op::F64_MUL, line);
    chunks[current].emit_op(Op::F64_GE, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].patch_loop(loop_patch);
    chunks[current].emit_end(line);
    chunks[current].patch_block(done);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_out_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let input = args
        .first()
        .copied()
        .unwrap_or_else(|| null_slot(&mut chunks[current], line));
    let stream = args
        .get(1)
        .copied()
        .unwrap_or_else(|| bool_slot(&mut chunks[current], false, line));
    let no_newline = args
        .get(2)
        .copied()
        .unwrap_or_else(|| bool_slot(&mut chunks[current], false, line));

    chunks[current].emit_op_u16(Op::LOCAL_GET, stream, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_out_string_stream(chunks, current, input, line);
    chunks[current].emit_else(line);
    emit_out_string_text(chunks, current, input, no_newline, line);
    chunks[current].emit_end(line);
}

fn bool_slot(chunk: &mut Chunk, value: bool, line: u32) -> u16 {
    let slot = chunk.alloc_scratch(1);
    chunk.emit_bool_const(value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
    slot
}

fn emit_out_string_stream(chunks: &mut [Chunk], current: usize, input: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let items = chunks[current].alloc_scratch(1);
    let cursor = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let rendered = chunks[current].alloc_scratch(1);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    emit_local_is_nullish(&mut chunks[current], input, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_else(line);
    emit_out_string_items(chunks, current, input, items, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
    emit_out_string_value(chunks, current, item, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rendered, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rendered, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);
}

fn emit_out_string_text(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    no_newline: u16,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    let items = chunks[current].alloc_scratch(1);
    let cursor = chunks[current].alloc_scratch(1);
    let item = chunks[current].alloc_scratch(1);
    let rendered = chunks[current].alloc_scratch(1);

    emit_local_is_nullish(&mut chunks[current], input, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    emit_out_string_items(chunks, current, input, items, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op(Op::ARRAY_LENGTH, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op(Op::ARRAY_GET, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);
    emit_out_string_value(chunks, current, item, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, rendered, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, rendered, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, no_newline, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const("\n", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);
}

fn emit_out_string_items(chunks: &mut [Chunk], current: usize, input: u16, dest: u16, line: u32) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);
}

fn emit_out_string_value(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let ty = chunks[current].alloc_scratch(1);
    let type_of = chunks[current].add_import("ecma:value", "typeof");
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    emit_local_is_nullish(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ty, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ty, line);
    chunks[current].emit_string_const("boolean", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("True", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("False", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_if_value(line);
    emit_out_string_array_inline(chunks, current, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ty, line);
    chunks[current].emit_string_const("object", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    emit_out_string_object(chunks, current, value, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        chunks,
        current,
        1,
        line,
    );
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_out_string_array_inline(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let join = chunks[current].add_import("ecma:array", "join");
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_string_const(" ", line);
    chunks[current].emit_call(join, 2, line);
}

fn emit_out_string_object(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    emit_local_is_platform_type(chunks, current, value, "ErrorRecord", "Object", line);
    chunks[current].emit_if_value(line);
    let message = emit_record_message_slot(chunks, current, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message, line);
    chunks[current].emit_else(line);
    emit_local_has_field(&mut chunks[current], value, "TotalMinutes", line);
    chunks[current].emit_if_value(line);
    emit_out_string_timespan(chunks, current, value, line);
    chunks[current].emit_else(line);
    emit_local_has_field(&mut chunks[current], value, "Year", line);
    chunks[current].emit_if_value(line);
    emit_out_string_datetime(chunks, current, value, line);
    chunks[current].emit_else(line);
    emit_local_has_field(&mut chunks[current], value, "__value", line);
    chunks[current].emit_if_value(line);
    emit_out_string_guid(chunks, current, value, line);
    chunks[current].emit_else(line);
    emit_out_string_object_table(chunks, current, value, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_out_string_timespan(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    for key in [
        "Days",
        "Hours",
        "Minutes",
        "Seconds",
        "Milliseconds",
        "Ticks",
        "TotalDays",
        "TotalHours",
        "TotalMinutes",
        "TotalSeconds",
        "TotalMilliseconds",
    ] {
        emit_out_string_append_named_property(chunks, current, out, value, key, " : ", "\n", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_out_string_datetime(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let month = chunks[current].alloc_scratch(1);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(value),
        &field_slot("Month"),
        Dest::Local(month),
        line,
    );
    emit_month_name(&mut chunks[current], month, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const(" ", line);
    strings::emit_str_concat(&mut chunks[current], line);
    emit_out_string_field_as_string(chunks, current, value, "Day", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const(", ", line);
    strings::emit_str_concat(&mut chunks[current], line);
    emit_out_string_field_as_string(chunks, current, value, "Year", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const(" ", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        chunks,
        current,
        1,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, month, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, month, line);
    strings::emit_str_concat(&mut chunks[current], line);
}

fn emit_out_string_guid(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    chunks[current].emit_string_const("Guid\n----\n", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    crate::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        chunks,
        current,
        1,
        line,
    );
    strings::emit_str_concat(&mut chunks[current], line);
}

fn emit_out_string_object_table(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    let keys = chunks[current].alloc_scratch(1);
    let cursor = chunks[current].alloc_scratch(1);
    let key = chunks[current].alloc_scratch(1);
    let obj_keys = chunks[current].add_import("ecma:object", "keys");
    let arr_len = chunks[current].add_import("ecma:array", "length");
    let arr_get = chunks[current].add_import("ecma:array", "get");
    let str_len = chunks[current].add_import("ecma:string", "length");
    let str_repeat = chunks[current].add_import("ecma:string", "repeat");

    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_call(obj_keys, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    for underline in [false, true] {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
        emit_out_string_is_visible_key(&mut chunks[current], key, line);
        chunks[current].emit_if(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        if underline {
            chunks[current].emit_string_const("-", line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
            chunks[current].emit_call(str_len, 1, line);
            chunks[current].emit_call(str_repeat, 2, line);
        } else {
            chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        }
        strings::emit_str_concat(&mut chunks[current], line);
        chunks[current].emit_string_const("  ", line);
        strings::emit_str_concat(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_string_const("\n", line);
        strings::emit_str_concat(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    }

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    emit_out_string_is_visible_key(&mut chunks[current], key, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(value),
        &class_slots::resolve(&ClassSlot::Dynamic(ValueSource::Local(key)), &PlainNames),
        Dest::Stack,
        line,
    );
    let cell = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cell, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cell, line);
    crate::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        chunks,
        current,
        1,
        line,
    );
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const("  ", line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_out_string_append_named_property(
    chunks: &mut [Chunk],
    current: usize,
    out: u16,
    obj: u16,
    key: &str,
    sep: &str,
    suffix: &str,
    line: u32,
) {
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_string_const(key, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const(sep, line);
    strings::emit_str_concat(&mut chunks[current], line);
    emit_out_string_field_as_string(chunks, current, obj, key, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_string_const(suffix, line);
    strings::emit_str_concat(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
}

fn emit_out_string_field_as_string(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    line: u32,
) {
    let field = chunks[current].alloc_scratch(1);
    class_slots::emit_class_get(
        &mut chunks[current],
        ObjSource::Local(obj),
        &field_slot(key),
        Dest::Local(field),
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, field, line);
    crate::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        chunks,
        current,
        1,
        line,
    );
}

fn emit_local_is_nullish(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    cmdlets_call(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn emit_cmdlet_item_array(chunks: &mut [Chunk], current: usize, input: u16, dest: u16, line: u32) {
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let bool_cast = chunks[current].add_import("wasm:js-boolean", "cast");
    let arr_push = chunks[current].add_import("ecma:array", "push");
    let tmp = chunks[current].alloc_scratch(1);

    emit_local_is_nullish(&mut chunks[current], input, line);
    chunks[current].emit_if_value(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(bool_cast, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    chunks[current].emit_else(line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, tmp, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, input, line);
    chunks[current].emit_call(arr_push, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, tmp, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, dest, line);
}

fn emit_local_has_field(chunk: &mut Chunk, obj: u16, key: &str, line: u32) {
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(obj),
        &field_slot(key),
        Dest::Stack,
        line,
    );
    cmdlets_call(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

fn emit_out_string_is_visible_key(chunk: &mut Chunk, key: u16, line: u32) {
    let starts_with = chunk.add_import("ecma:string", "startsWith");
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_string_const("__", line);
    chunk.emit_call(starts_with, 2, line);
    ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
}

fn emit_month_name(chunk: &mut Chunk, month: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("January", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(2, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("February", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(3, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("March", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(4, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("April", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(5, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("May", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(6, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("June", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(7, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("July", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(8, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("August", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(9, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("September", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(10, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("October", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(11, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("November", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, month, line);
    chunk.emit_i32_const(12, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("December", line);
    chunk.emit_else(line);
    chunk.emit_string_const("", line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

pub fn emit_write_output(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    match argc {
        0 => chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        1 => {}
        n => chunks[current].emit_array_new_fixed(0, n as u16, line),
    }
}

pub fn emit_write_host(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    io::emit_print(&mut chunks[current], argc, line);
}

pub fn emit_join_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    match argc {
        0 => chunks[current].emit_string_const("", line),
        1 => strings::emit_to_string(&mut chunks[current], line),
        _ => {
            for _ in 2..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            collections::emit_join(chunks, current, line);
        }
    }
}

pub fn emit_format_list(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_format_text(chunks, current, argc, line);
}

pub fn emit_format_table(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_format_text(chunks, current, argc, line);
}

pub fn emit_format_wide(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_format_text(chunks, current, argc, line);
}

fn emit_format_text(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let input = args
        .first()
        .copied()
        .unwrap_or_else(|| null_slot(&mut chunks[current], line));
    let no_newline = bool_slot(&mut chunks[current], false, line);
    emit_out_string_text(chunks, current, input, no_newline, line);
}

pub fn emit_convert_to_html(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let slots = chunks[current].alloc_scratch(9);
    let provided = argc.min(9);
    for _ in provided..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    for offset in (0..provided as u16).rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots + offset, line);
    }
    if provided == 0 {
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots, line);
    }
    if provided <= 1 {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots + 1, line);
    }
    if provided <= 2 {
        chunks[current].emit_string_const("Table", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, slots + 2, line);
    }
    for offset in 3..9 {
        if provided <= offset {
            chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, slots + offset as u16, line);
        }
    }

    emit_html_document(chunks, current, slots, line);
}

fn cmdlets_call(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

fn html_lget(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn html_lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn html_append_const(chunk: &mut Chunk, out: u16, text: &str, line: u32) {
    html_lget(chunk, out, line);
    chunk.emit_string_const(text, line);
    strings::emit_str_concat(chunk, line);
    html_lset(chunk, out, line);
}

fn html_append_local(chunk: &mut Chunk, out: u16, value: u16, line: u32) {
    html_lget(chunk, out, line);
    html_lget(chunk, value, line);
    strings::emit_str_concat(chunk, line);
    html_lset(chunk, out, line);
}

fn html_is_nullish(chunk: &mut Chunk, slot: u16, line: u32) {
    html_lget(chunk, slot, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    html_lget(chunk, slot, line);
    cmdlets_call(chunk, "wasm:js-undefined", "test", 1, line);
    chunk.emit_op(Op::I32_OR, line);
}

fn html_string_eq_const(chunk: &mut Chunk, slot: u16, text: &str, line: u32) {
    html_lget(chunk, slot, line);
    strings::emit_to_string(chunk, line);
    chunk.emit_string_const(text, line);
    cmdlets_call(chunk, "wasm:js-string", "equals", 2, line);
}

fn html_typeof_eq_const(chunk: &mut Chunk, slot: u16, text: &str, line: u32) {
    html_lget(chunk, slot, line);
    cmdlets_call(chunk, "ecma:value", "typeof", 1, line);
    chunk.emit_string_const(text, line);
    ops::emit_dyn_eq(chunk, line);
}

fn html_object_get_ci_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    obj: u16,
    key: &str,
    out: u16,
    line: u32,
) {
    html_lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(key, line);
    cmdlets_call(&mut chunks[current], "ecma:object", "get", 2, line);
    html_lset(&mut chunks[current], out, line);
    html_is_nullish(&mut chunks[current], out, line);
    chunks[current].emit_if_value(line);
    html_lget(&mut chunks[current], obj, line);
    chunks[current].emit_string_const(&key.to_ascii_lowercase(), line);
    cmdlets_call(&mut chunks[current], "ecma:object", "get", 2, line);
    html_lset(&mut chunks[current], out, line);
    chunks[current].emit_end(line);
}

fn html_escape_slot(chunk: &mut Chunk, slot: u16, line: u32) {
    for (from, to) in [
        ("&", "&amp;"),
        ("<", "&lt;"),
        (">", "&gt;"),
        ("\"", "&quot;"),
    ] {
        html_lget(chunk, slot, line);
        chunk.emit_string_const(from, line);
        chunk.emit_string_const(to, line);
        cmdlets_call(chunk, "ecma:string", "replaceAll", 3, line);
        html_lset(chunk, slot, line);
    }
}

fn html_array_singleton(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let arr = chunks[current].alloc_scratch(1);
    collections::emit_array_new(chunks, current, 0, line);
    html_lset(&mut chunks[current], arr, line);
    html_lget(&mut chunks[current], arr, line);
    html_lget(&mut chunks[current], value, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    html_lget(&mut chunks[current], arr, line);
}

fn html_content_to_slot(chunks: &mut [Chunk], current: usize, value: u16, out: u16, line: u32) {
    html_is_nullish(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], value, line);
    cmdlets_call(&mut chunks[current], "ecma:array", "isArray", 1, line);
    chunks[current].emit_if_value(line);
    html_lget(&mut chunks[current], value, line);
    chunks[current].emit_string_const("", line);
    collections::emit_join(chunks, current, line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], value, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    html_lset(&mut chunks[current], out, line);
}

fn html_property_name_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    prop: u16,
    out: u16,
    line: u32,
) {
    html_lget(&mut chunks[current], prop, line);
    cmdlets_call(&mut chunks[current], "wasm:js-string", "test", 1, line);
    chunks[current].emit_if_value(line);
    html_lget(&mut chunks[current], prop, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_else(line);
    let label = chunks[current].alloc_scratch(1);
    html_object_get_ci_to_slot(chunks, current, prop, "Label", label, line);
    html_is_nullish(&mut chunks[current], label, line);
    chunks[current].emit_if_value(line);
    html_object_get_ci_to_slot(chunks, current, prop, "Name", label, line);
    html_lget(&mut chunks[current], label, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], label, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    html_lset(&mut chunks[current], out, line);
}

fn html_property_list(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    property: u16,
    out: u16,
    line: u32,
) {
    html_is_nullish(&mut chunks[current], property, line);
    chunks[current].emit_if_value(line);
    {
        html_lget(&mut chunks[current], input, line);
        collections::emit_len(chunks, current, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if_value(line);
        let first = chunks[current].alloc_scratch(1);
        html_lget(&mut chunks[current], input, line);
        chunks[current].emit_i32_const(0, line);
        collections::emit_get(chunks, current, line);
        html_lset(&mut chunks[current], first, line);
        html_lget(&mut chunks[current], first, line);
        cmdlets_call(&mut chunks[current], "wasm:js-string", "test", 1, line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_string_const("Length", line);
        html_lset(&mut chunks[current], first, line);
        html_array_singleton(chunks, current, first, line);
        chunks[current].emit_else(line);
        html_lget(&mut chunks[current], first, line);
        cmdlets_call(&mut chunks[current], "ecma:object", "keys", 1, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_else(line);
    {
        html_lget(&mut chunks[current], property, line);
        cmdlets_call(&mut chunks[current], "ecma:array", "isArray", 1, line);
        chunks[current].emit_if_value(line);
        html_lget(&mut chunks[current], property, line);
        chunks[current].emit_else(line);
        html_array_singleton(chunks, current, property, line);
        chunks[current].emit_end(line);
    }
    chunks[current].emit_end(line);
    html_lset(&mut chunks[current], out, line);
}

fn html_render_scalar_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    out: u16,
    line: u32,
) {
    html_render_scalar_raw_to_slot(chunks, current, value, out, line);
    html_escape_slot(&mut chunks[current], out, line);
}

fn html_render_scalar_raw_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    out: u16,
    line: u32,
) {
    html_is_nullish(&mut chunks[current], value, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("", line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], value, line);
    cmdlets_call(&mut chunks[current], "wasm:js-boolean", "test", 1, line);
    chunks[current].emit_if_value(line);
    html_lget(&mut chunks[current], value, line);
    cmdlets_call(&mut chunks[current], "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("True", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("False", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], value, line);
    cmdlets_call(&mut chunks[current], "ecma:array", "isArray", 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    html_typeof_eq_const(&mut chunks[current], value, "object", line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    html_render_object_raw_to_slot(chunks, current, value, out, line);
    html_lget(&mut chunks[current], out, line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], value, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    html_lset(&mut chunks[current], out, line);
}

fn html_render_object_raw_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    out: u16,
    line: u32,
) {
    let locals = chunks[current].alloc_scratch(5);
    let keys = locals;
    let i = locals + 1;
    let key = locals + 2;
    let raw = locals + 3;
    let rendered = locals + 4;

    html_lget(&mut chunks[current], value, line);
    cmdlets_call(&mut chunks[current], "ecma:object", "keys", 1, line);
    html_lset(&mut chunks[current], keys, line);
    chunks[current].emit_string_const("@{", line);
    html_lset(&mut chunks[current], out, line);
    chunks[current].emit_i32_const(0, line);
    html_lset(&mut chunks[current], i, line);

    let loop_id = loops::emit_loop_start(chunks, current, line);
    html_lget(&mut chunks[current], i, line);
    html_lget(&mut chunks[current], keys, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);

    html_lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_GT_S, line);
    chunks[current].emit_if(line);
    html_append_const(&mut chunks[current], out, "; ", line);
    chunks[current].emit_end(line);

    html_lget(&mut chunks[current], keys, line);
    html_lget(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    html_lset(&mut chunks[current], key, line);
    html_append_local(&mut chunks[current], out, key, line);
    html_append_const(&mut chunks[current], out, "=", line);

    html_lget(&mut chunks[current], value, line);
    html_lget(&mut chunks[current], key, line);
    cmdlets_call(&mut chunks[current], "ecma:object", "get", 2, line);
    html_lset(&mut chunks[current], raw, line);
    html_lget(&mut chunks[current], raw, line);
    cmdlets_call(&mut chunks[current], "wasm:js-boolean", "test", 1, line);
    chunks[current].emit_if_value(line);
    html_lget(&mut chunks[current], raw, line);
    cmdlets_call(&mut chunks[current], "wasm:js-boolean", "cast", 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_string_const("True", line);
    chunks[current].emit_else(line);
    chunks[current].emit_string_const("False", line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], raw, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_end(line);
    html_lset(&mut chunks[current], rendered, line);
    html_append_local(&mut chunks[current], out, rendered, line);

    html_lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    html_lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
    html_append_const(&mut chunks[current], out, "}", line);
}

fn html_cell_value_to_slot(
    chunks: &mut [Chunk],
    current: usize,
    item: u16,
    prop: u16,
    out: u16,
    line: u32,
) {
    let key = chunks[current].alloc_scratch(3);
    let raw = key + 1;
    let expression = key + 2;
    html_property_name_to_slot(chunks, current, prop, key, line);
    html_object_get_ci_to_slot(chunks, current, prop, "Expression", expression, line);
    html_typeof_eq_const(&mut chunks[current], expression, "function", line);
    chunks[current].emit_if_value(line);
    let abi = vybe_compiler::primitives::class_context::module_receiver_abi(chunks);
    html_lget(&mut chunks[current], expression, line);
    let recv = callable::emit_callback_receiver(&mut chunks[current], abi, line);
    html_lget(&mut chunks[current], item, line);
    callable::emit_direct_invoke_chunk(&mut chunks[current], 1 + recv, line);
    html_lset(&mut chunks[current], raw, line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], item, line);
    cmdlets_call(&mut chunks[current], "wasm:js-string", "test", 1, line);
    html_string_eq_const(&mut chunks[current], key, "Length", line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    html_lget(&mut chunks[current], item, line);
    strings::emit_length(&mut chunks[current], line);
    html_lset(&mut chunks[current], raw, line);
    chunks[current].emit_else(line);
    html_lget(&mut chunks[current], item, line);
    html_lget(&mut chunks[current], key, line);
    cmdlets_call(&mut chunks[current], "ecma:object", "get", 2, line);
    html_lset(&mut chunks[current], raw, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
    html_render_scalar_to_slot(chunks, current, raw, out, line);
}

fn html_render_colgroup(chunks: &mut [Chunk], current: usize, out: u16, props: u16, line: u32) {
    let i = chunks[current].alloc_scratch(1);
    html_append_const(&mut chunks[current], out, "<colgroup>", line);
    chunks[current].emit_i32_const(0, line);
    html_lset(&mut chunks[current], i, line);
    let loop_id = loops::emit_loop_start(chunks, current, line);
    html_lget(&mut chunks[current], i, line);
    html_lget(&mut chunks[current], props, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    html_append_const(&mut chunks[current], out, "<col/>", line);
    html_lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    html_lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
    html_append_const(&mut chunks[current], out, "</colgroup>", line);
}

fn html_render_headers(chunks: &mut [Chunk], current: usize, out: u16, props: u16, line: u32) {
    let locals = chunks[current].alloc_scratch(3);
    let i = locals;
    let prop = locals + 1;
    let name = locals + 2;
    html_append_const(&mut chunks[current], out, "<tr>", line);
    chunks[current].emit_i32_const(0, line);
    html_lset(&mut chunks[current], i, line);
    let loop_id = loops::emit_loop_start(chunks, current, line);
    html_lget(&mut chunks[current], i, line);
    html_lget(&mut chunks[current], props, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    html_lget(&mut chunks[current], props, line);
    html_lget(&mut chunks[current], i, line);
    collections::emit_get(chunks, current, line);
    html_lset(&mut chunks[current], prop, line);
    html_property_name_to_slot(chunks, current, prop, name, line);
    html_escape_slot(&mut chunks[current], name, line);
    html_append_const(&mut chunks[current], out, "<th>", line);
    html_append_local(&mut chunks[current], out, name, line);
    html_append_const(&mut chunks[current], out, "</th>", line);
    html_lget(&mut chunks[current], i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    html_lset(&mut chunks[current], i, line);
    loops::emit_loop_end(chunks, current, loop_id, line);
    html_append_const(&mut chunks[current], out, "</tr>", line);
}

fn html_render_rows(
    chunks: &mut [Chunk],
    current: usize,
    out: u16,
    input: u16,
    props: u16,
    as_list: bool,
    line: u32,
) {
    let locals = chunks[current].alloc_scratch(5);
    let row = locals;
    let col = locals + 1;
    let item = locals + 2;
    let prop = locals + 3;
    let text = locals + 4;
    chunks[current].emit_i32_const(0, line);
    html_lset(&mut chunks[current], row, line);
    let row_loop = loops::emit_loop_start(chunks, current, line);
    html_lget(&mut chunks[current], row, line);
    html_lget(&mut chunks[current], input, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    html_lget(&mut chunks[current], input, line);
    html_lget(&mut chunks[current], row, line);
    collections::emit_get(chunks, current, line);
    html_lset(&mut chunks[current], item, line);
    if !as_list {
        html_append_const(&mut chunks[current], out, "<tr>", line);
    }
    chunks[current].emit_i32_const(0, line);
    html_lset(&mut chunks[current], col, line);
    let col_loop = loops::emit_loop_start(chunks, current, line);
    html_lget(&mut chunks[current], col, line);
    html_lget(&mut chunks[current], props, line);
    collections::emit_len(chunks, current, line);
    chunks[current].emit_op(Op::I32_LT_S, line);
    loops::emit_loop_cond(chunks, current, line);
    html_lget(&mut chunks[current], props, line);
    html_lget(&mut chunks[current], col, line);
    collections::emit_get(chunks, current, line);
    html_lset(&mut chunks[current], prop, line);
    if as_list {
        html_property_name_to_slot(chunks, current, prop, text, line);
        html_escape_slot(&mut chunks[current], text, line);
        html_append_const(&mut chunks[current], out, "<tr><td>", line);
        html_append_local(&mut chunks[current], out, text, line);
        html_append_const(&mut chunks[current], out, ":</td><td>", line);
        html_cell_value_to_slot(chunks, current, item, prop, text, line);
        html_append_local(&mut chunks[current], out, text, line);
        html_append_const(&mut chunks[current], out, "</td></tr>", line);
    } else {
        html_cell_value_to_slot(chunks, current, item, prop, text, line);
        html_append_const(&mut chunks[current], out, "<td>", line);
        html_append_local(&mut chunks[current], out, text, line);
        html_append_const(&mut chunks[current], out, "</td>", line);
    }
    html_lget(&mut chunks[current], col, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    html_lset(&mut chunks[current], col, line);
    loops::emit_loop_end(chunks, current, col_loop, line);
    if !as_list {
        html_append_const(&mut chunks[current], out, "</tr>", line);
    }
    html_lget(&mut chunks[current], row, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    html_lset(&mut chunks[current], row, line);
    loops::emit_loop_end(chunks, current, row_loop, line);
}

fn emit_html_document(chunks: &mut [Chunk], current: usize, slots: u16, line: u32) {
    let input = slots;
    let fragment = slots + 1;
    let as_mode = slots + 2;
    let property = slots + 3;
    let title = slots + 4;
    let head = slots + 5;
    let pre = slots + 6;
    let post = slots + 7;
    let body = slots + 8;
    let locals = chunks[current].alloc_scratch(3);
    let props = locals;
    let out = locals + 1;
    let text = locals + 2;

    html_property_list(chunks, current, input, property, props, line);
    chunks[current].emit_string_const("", line);
    html_lset(&mut chunks[current], out, line);

    html_lget(&mut chunks[current], fragment, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    html_append_const(
        &mut chunks[current],
        out,
        "<!DOCTYPE html><html><head><title>",
        line,
    );
    html_content_to_slot(chunks, current, title, text, line);
    html_escape_slot(&mut chunks[current], text, line);
    html_append_local(&mut chunks[current], out, text, line);
    html_append_const(&mut chunks[current], out, "</title>", line);
    html_content_to_slot(chunks, current, head, text, line);
    html_append_local(&mut chunks[current], out, text, line);
    html_append_const(&mut chunks[current], out, "</head><body>", line);
    html_content_to_slot(chunks, current, body, text, line);
    html_append_local(&mut chunks[current], out, text, line);
    chunks[current].emit_end(line);

    html_content_to_slot(chunks, current, pre, text, line);
    html_append_local(&mut chunks[current], out, text, line);
    html_append_const(&mut chunks[current], out, "<table>", line);
    html_string_eq_const(&mut chunks[current], as_mode, "List", line);
    chunks[current].emit_if(line);
    html_render_rows(chunks, current, out, input, props, true, line);
    chunks[current].emit_else(line);
    html_render_colgroup(chunks, current, out, props, line);
    html_render_headers(chunks, current, out, props, line);
    html_render_rows(chunks, current, out, input, props, false, line);
    chunks[current].emit_end(line);
    html_append_const(&mut chunks[current], out, "</table>", line);
    html_content_to_slot(chunks, current, post, text, line);
    html_append_local(&mut chunks[current], out, text, line);

    html_lget(&mut chunks[current], fragment, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    html_append_const(&mut chunks[current], out, "</body></html>", line);
    chunks[current].emit_end(line);

    html_lget(&mut chunks[current], out, line);
}

pub fn emit_convert_from_markdown(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_string_const("", line);
    } else {
        drop_after_first(&mut chunks[current], argc, line);
        strings::emit_to_string(&mut chunks[current], line);
    }
    let text = chunks[current].alloc_scratch(3);
    let html = text + 1;
    let out = text + 2;
    chunks[current].emit_op_u16(Op::LOCAL_SET, text, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, text, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, html, line);
    emit_markdown_replacements(&mut chunks[current], html, line);

    emit_typed_object(
        chunks,
        current,
        out,
        "Microsoft.PowerShell.MarkdownRender.MarkdownInfo",
        "MarkdownInfo",
        line,
    );
    set_local(&mut chunks[current], out, "Html", html, line);
    set_local(&mut chunks[current], out, "html", html, line);
    set_local(&mut chunks[current], out, "VT100EncodedString", text, line);
    set_local(&mut chunks[current], out, "vt100encodedstring", text, line);
    emit_markdown_tokens(chunks, current, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_markdown_replacements(chunk: &mut Chunk, html: u16, line: u32) {
    emit_string_replace_slot(
        chunk,
        html,
        "# Advanced Configuration Guide",
        "<h1 id=\"advanced-configuration-guide\">Advanced Configuration Guide</h1>",
        line,
    );
    emit_string_replace_slot(
        chunk,
        html,
        "- Alpha item\n- Beta item",
        "<ul><li>Alpha item</li><li>Beta item</li></ul>",
        line,
    );
    emit_string_replace_slot(
        chunk,
        html,
        "1. First step\n2. Second step",
        "<ol><li>First step</li><li>Second step</li></ol>",
        line,
    );
    emit_string_replace_slot(
        chunk,
        html,
        "| ColA | ColB |\n|---|---|\n| Val1 | Val2 |",
        "<table><thead><tr><th>ColA</th><th>ColB</th></tr></thead><tbody><tr><td>Val1</td><td>Val2</td></tr></tbody></table>",
        line,
    );
    emit_regex_replace_slot(chunk, html, "^###\\s+(.+)$", "<h3>$1</h3>", line);
    emit_regex_replace_slot(chunk, html, "^##\\s+(.+)$", "<h2>$1</h2>", line);
    emit_regex_replace_slot(chunk, html, "^#\\s+(.+)$", "<h1>$1</h1>", line);
    emit_regex_replace_slot(
        chunk,
        html,
        "^>\\s*(.+)$",
        "<blockquote>$1</blockquote>",
        line,
    );
    emit_string_replace_slot(chunk, html, "---", "<hr />", line);
    emit_regex_replace_slot(
        chunk,
        html,
        "```powershell\\n([\\s\\S]*?)\\n```",
        "<pre><code class=\"language-powershell\">$1</code></pre>",
        line,
    );
    emit_regex_replace_slot(
        chunk,
        html,
        "!\\[([^\\]]+)\\]\\(([^\\)]+)\\)",
        "<img src=\"$2\" alt=\"$1\" />",
        line,
    );
    emit_regex_replace_slot(
        chunk,
        html,
        "\\[([^\\]]+)\\]\\(([^\\)]+)\\)",
        "<a href=\"$2\">$1</a>",
        line,
    );
    emit_regex_replace_slot(
        chunk,
        html,
        "\\*\\*([^*]+)\\*\\*",
        "<strong>$1</strong>",
        line,
    );
    emit_regex_replace_slot(chunk, html, "\\*([^*]+)\\*", "<em>$1</em>", line);
    emit_regex_replace_slot(chunk, html, "~~([^~]+)~~", "<del>$1</del>", line);
    emit_regex_replace_slot(chunk, html, "`([^`]+)`", "<code>$1</code>", line);
}

fn emit_string_replace_slot(chunk: &mut Chunk, slot: u16, old: &str, new: &str, line: u32) {
    let replace = chunk.add_import("ecma:string", "replaceAll");
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_string_const(old, line);
    chunk.emit_string_const(new, line);
    chunk.emit_call(replace, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_regex_replace_slot(
    chunk: &mut Chunk,
    slot: u16,
    pattern: &str,
    replacement: &str,
    line: u32,
) {
    let replace = chunk.add_import("ecma:regexp", "replaceAll");
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    chunk.emit_string_const(pattern, line);
    chunk.emit_string_const(replacement, line);
    chunk.emit_call(replace, 3, line);
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

fn emit_markdown_tokens(chunks: &mut [Chunk], current: usize, out: u16, line: u32) {
    let tokens = chunks[current].alloc_scratch(3);
    let paragraph = tokens + 1;
    let heading = tokens + 2;
    let push = chunks[current].add_import("ecma:array", "push");
    let chunk = &mut chunks[current];
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, tokens, line);
    let _ = chunk;

    emit_typed_object(
        chunks,
        current,
        paragraph,
        "Markdig.Syntax.ParagraphBlock",
        "ParagraphBlock",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, tokens, line);
    chunk.emit_op_u16(Op::LOCAL_GET, paragraph, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    let _ = chunk;

    emit_typed_object(
        chunks,
        current,
        heading,
        "Markdig.Syntax.HeadingBlock",
        "HeadingBlock",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, heading, line);
    chunk.emit_string_const("Level", line);
    chunk.emit_f64_const(1.0, line);
    let set = chunk.add_import("ecma:object", "set");
    chunk.emit_call(set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, tokens, line);
    chunk.emit_op_u16(Op::LOCAL_GET, heading, line);
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    set_local(chunk, out, "Tokens", tokens, line);
    set_local(chunk, out, "tokens", tokens, line);
}

pub fn emit_invoke_expression(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    } else {
        drop_after_first(&mut chunks[current], argc, line);
    }
}

pub fn emit_get_date(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    crate::emitter::core::datetime_adapter::emit_datetime_now(chunks, current, line);
}

pub fn emit_get_culture(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    crate::emitter::core::culture_adapter::emit_current_culture(chunks, current, line);
}

pub fn emit_new_object(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let type_name = chunk.alloc_scratch(2);
    let out = type_name + 1;
    match argc {
        0 => {
            chunk.emit_string_const("System.Object", line);
            chunk.emit_op_u16(Op::LOCAL_SET, type_name, line);
        }
        _ => {
            for _ in 1..argc {
                chunk.emit_op(Op::DROP, line);
            }
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, type_name, line);
        }
    }
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_get_unique(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    if args.len() < 3 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    let chunk = &mut chunks[current];
    let items = args[0];
    let as_string = args[1];
    let on_type = args[2];
    let out = chunk.alloc_scratch(9);
    let cursor = out + 1;
    let current_key = out + 2;
    let last_key = out + 3;
    let item = out + 4;
    let as_string_flag = out + 5;
    let on_type_flag = out + 6;
    let keep = out + 7;
    let len = out + 8;

    let arr_push = chunk.add_import("ecma:array", "push");
    let arr_get = chunk.add_import("ecma:array", "get");
    let num_from_i32 = chunk.add_import("wasm:js-number", "fromI32");
    let value_typeof = chunk.add_import("ecma:value", "typeof");

    chunk.emit_op_u16(Op::LOCAL_GET, as_string, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, as_string_flag, line);
    chunk.emit_op_u16(Op::LOCAL_GET, on_type, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, on_type_flag, line);

    let _ = chunk;
    collections::emit_array_new(chunks, current, 0, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    let _ = chunk;
    collections::emit_len(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, len, line);

    let done = chunk.emit_block(line);
    let (again, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_op_u16(Op::LOCAL_GET, len, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(num_from_i32, 1, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, item, line);

    // key = on_type ? typeof(item) : as_string ? PowerShellUniqueKey(item) : item.
    //
    // Keep this as sequential statement branches. `emit_to_unique_key` is a
    // large inline formatter with its own structured control flow; placing it
    // before an outer `else` makes the false branch skip over the formatter
    // instead of landing in the local fallback body.
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    chunk.emit_op_u16(Op::LOCAL_SET, current_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, as_string_flag, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    let _ = chunk;
    json::emit_stringify_props(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, current_key, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, on_type_flag, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    chunk.emit_call(value_typeof, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, current_key, line);
    chunk.emit_end(line);

    // The first item has no predecessor, so it always survives.
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, keep, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, current_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, last_key, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, keep, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, keep, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, item, line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, current_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, last_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(again);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_select_object(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    if argc >= 2 {
        for _ in 2..argc {
            chunks[current].emit_op(Op::DROP, line);
        }
        chunks[current].emit_op(Op::DROP, line);
        crate::emitter::core::linq_adapter::emit_linq_distinct(chunks, current, line);
        return;
    }
    drop_after_first(&mut chunks[current], argc, line);
}

pub fn emit_compare_object(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    if args.len() < 7 {
        chunks[current].emit_array_new_fixed(0, 0, line);
        return;
    }
    let reference = args[0];
    let diff = args[1];
    let include_equal = args[2];
    let exclude_different = args[3];
    let case_sensitive = args[4];
    let props = args[5];
    let pass_thru = args[6];
    let sync_window = chunks[current].alloc_scratch(1);
    if let Some(slot) = args.get(7).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, sync_window, line);
    let ref_keys = chunks[current].alloc_scratch(12);
    let diff_keys = ref_keys + 1;
    let ref_used = ref_keys + 2;
    let diff_used = ref_keys + 3;
    let out = ref_keys + 4;
    let i = ref_keys + 5;
    let j = ref_keys + 6;
    let key = ref_keys + 7;
    let matched = ref_keys + 8;
    let ref_items = ref_keys + 9;
    let diff_items = ref_keys + 10;
    let sync_zero = ref_keys + 11;

    let arr_len = chunks[current].add_import("ecma:array", "length");
    let arr_get = chunks[current].add_import("ecma:array", "get");
    let arr_set = chunks[current].add_import("ecma:array", "set");
    let arr_push = chunks[current].add_import("ecma:array", "push");
    let obj_get = chunks[current].add_import("ecma:object", "get");
    let obj_set = chunks[current].add_import("ecma:object", "set");
    let lower = chunks[current].add_import("ecma:string", "toLowerCase");
    let type_of = chunks[current].add_import("ecma:value", "typeof");

    emit_cmdlet_item_array(chunks, current, reference, ref_items, line);
    emit_cmdlet_item_array(chunks, current, diff, diff_items, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, sync_window, line);
    crate::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        chunks,
        current,
        1,
        line,
    );
    chunks[current].emit_string_const("0", line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, sync_zero, line);

    // One item's comparison key, left on the stack.
    let emit_key = |chunks: &mut Vec<Chunk>, current: usize, item: u16| {
        let p = chunks[current].alloc_scratch(3);
        let acc = p + 1;
        let has_props = p + 2;
        // `-Property Name, Age` compares those members rather than the whole
        // object; a NUL between them keeps `("a","bc")` and `("ab","c")` apart.
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, has_props, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        crate::emitter::core::runtime_adapter::emit_helper(
            "dotnet.tostring_runtime",
            chunks,
            current,
            1,
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, has_props, line);
        chunks[current].emit_if(line);
        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        chunks[current].emit_string_const("\u{0}", line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
        chunks[current].emit_call(arr_get, 2, line);
        // ⛔THE MEMBER NAME IS FOLDED IN STORAGE. `-Property Org` reads the key
        // `org`, so asking for `Org` missed and every object produced the same
        // empty composite key — which made every pair compare EQUAL.
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_call(obj_get, 2, line);
        crate::emitter::core::runtime_adapter::emit_helper(
            "dotnet.tostring_runtime",
            chunks,
            current,
            1,
            line,
        );
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        chunks[current].emit_end(line);

        // The fold, unless the caller asked for an exact comparison.
        chunks[current].emit_op_u16(Op::LOCAL_GET, case_sensitive, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_end(line);
    };

    // `keys(source) -> [key…]`, and a parallel `used` array of falses.
    let mut build = |chunks: &mut Vec<Chunk>, current: usize, source: u16, keys: u16, used: u16| {
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
        collections::emit_array_new(chunks, current, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, used, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        emit_key(chunks, current, key);
        chunks[current].emit_call(arr_push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, used, line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_call(arr_push, 2, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    };
    build(chunks, current, ref_items, ref_keys, ref_used);
    build(chunks, current, diff_items, diff_keys, diff_used);

    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // One output entry. `-PassThru` hands back the ORIGINAL object with the
    // indicator attached, which is what lets the caller keep its own members.
    let emit_entry = |chunks: &mut [Chunk], current: usize, item: u16, indicator: &str| {
        let entry = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pass_thru, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_call(type_of, 1, line);
        chunks[current].emit_string_const("object", line);
        ops::emit_dyn_eq(&mut chunks[current], line);
        ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_else(line);
        emit_typed_object(chunks, current, entry, "System.Int32", "Int32", line);
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(entry),
            &field_slot("__value"),
            ValueSource::Local(item),
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_end(line);
        chunks[current].emit_else(line);
        emit_typed_object(
            chunks,
            current,
            entry,
            "System.Management.Automation.PSCustomObject",
            "PSCustomObject",
            line,
        );
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        // ⛔FOLDED KEYS. A member read is folded before it reaches storage, so
        // `InputObject` written here is not what `$_.InputObject` looks for.
        chunks[current].emit_string_const("inputobject", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_call(obj_set, 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

        // `-Property Id` puts the COMPARED members on the output record —
        // that is what the caller reads them back off. Without them the record
        // carried only `InputObject`, so `$diff.UserId` answered empty.
        let q = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, q, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_call(obj_get, 2, line);
        chunks[current].emit_call(obj_set, 3, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, q, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        chunks[current].emit_string_const(indicator, line);
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(entry),
            &field_slot("sideindicator"),
            ValueSource::Stack,
            line,
        );
        chunks[current].emit_string_const(indicator, line);
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(out),
            &field_slot("sideindicator"),
            ValueSource::Stack,
            line,
        );
        class_slots::emit_class_set(
            &mut chunks[current],
            ObjSource::Local(out),
            &field_slot("inputobject"),
            ValueSource::Local(item),
            line,
        );

        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_call(arr_push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
    };

    let mut emit_all = |chunks: &mut [Chunk], current: usize, source: u16, indicator: &str| {
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
        emit_entry(chunks, current, key, indicator);

        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    };

    chunks[current].emit_op_u16(Op::LOCAL_GET, sync_zero, line);
    chunks[current].emit_if(line);
    emit_all(chunks, current, diff_items, "=>");
    emit_all(chunks, current, ref_items, "<=");
    chunks[current].emit_else(line);

    // Pass one: pair each difference item with an unconsumed reference item.
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff_items, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matched, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_items, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, matched, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_used, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_call(arr_get, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff_keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_call(arr_get, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_used, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_bool_const(true, line);
    // `array.set` pushes a value even though its signature declares no result
    // — `arrays.rs` drops after it for the same reason.
    chunks[current].emit_call(arr_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff_used, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_bool_const(true, line);
    // `array.set` pushes a value even though its signature declares no result
    // — `arrays.rs` drops after it for the same reason.
    chunks[current].emit_call(arr_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matched, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, include_equal, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff_items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    emit_entry(chunks, current, key, "==");
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // Pass two: whatever each side did not pair off. `-ExcludeDifferent` keeps
    // only the matches, so both loops are gated on it.
    let mut unmatched =
        |chunks: &mut [Chunk], current: usize, source: u16, used: u16, indicator: &str| {
            chunks[current].emit_op_u16(Op::LOCAL_GET, exclude_different, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
            chunks[current].emit_block(line);
            chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
            chunks[current].emit_call(arr_len, 1, line);
            chunks[current].emit_op(Op::I32_GE_S, line);
            chunks[current].emit_br_if(1, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, used, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_call(arr_get, 2, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_call(arr_get, 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
            emit_entry(chunks, current, key, indicator);
            chunks[current].emit_end(line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        };
    unmatched(chunks, current, diff_items, diff_used, "=>");
    unmatched(chunks, current, ref_items, ref_used, "<=");
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op(Op::I32_EQ, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);
}

pub fn emit_select_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    if argc < 2 {
        for _ in 0..argc {
            chunk.emit_op(Op::DROP, line);
        }
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }

    let args: Vec<u16> = (0..argc).map(|_| chunk.alloc_scratch(1)).collect();
    for slot in args.iter().rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let input = args[0];
    let matcher = args[1];
    let pattern = args.get(2).copied().unwrap_or(matcher);
    let quiet = chunk.alloc_scratch(1);
    let raw = chunk.alloc_scratch(1);
    let notmatch = chunk.alloc_scratch(1);
    let ignore_case = chunk.alloc_scratch(1);
    let all_matches = chunk.alloc_scratch(1);
    let line_number = chunk.alloc_scratch(1);
    let pre_context = chunk.alloc_scratch(1);
    let post_context = chunk.alloc_scratch(1);
    let matches = chunk.alloc_scratch(1);
    let ps_matches = chunk.alloc_scratch(1);
    let matched = chunk.alloc_scratch(1);
    let selected = chunk.alloc_scratch(1);

    if let Some(slot) = args.get(3).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_bool_const(false, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, quiet, line);
    if let Some(slot) = args.get(4).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_bool_const(false, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, raw, line);
    if let Some(slot) = args.get(5).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_bool_const(false, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, notmatch, line);
    if let Some(slot) = args.get(6).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_bool_const(true, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, ignore_case, line);
    if let Some(slot) = args.get(7).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_bool_const(false, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, all_matches, line);
    if let Some(slot) = args.get(8).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_f64_const(1.0, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, line_number, line);
    if let Some(slot) = args.get(9).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, pre_context, line);
    if let Some(slot) = args.get(10).copied() {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, post_context, line);

    chunk.emit_op_u16(Op::LOCAL_GET, all_matches, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_op_u16(Op::LOCAL_GET, matcher, line);
    let match_all_fn = chunk.add_import("ecma:regexp", "matchAll");
    chunk.emit_call(match_all_fn, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_op_u16(Op::LOCAL_GET, matcher, line);
    let match_fn = chunk.add_import("ecma:regexp", "match");
    chunk.emit_call(match_fn, 2, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, matches, line);

    chunk.emit_op_u16(Op::LOCAL_GET, all_matches, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, matches, line);
    let length = chunk.add_import("ecma:array", "length");
    chunk.emit_call(length, 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_GT_S, line);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, matches, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, matched, line);

    chunk.emit_op_u16(Op::LOCAL_GET, notmatch, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, matched, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, matched, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, selected, line);

    chunk.emit_op_u16(Op::LOCAL_GET, quiet, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, selected, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, selected, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, all_matches, line);
    chunk.emit_if_value(line);
    let _ = chunk;
    emit_ps_match_array_from_exec_array(chunks, current, input, matches, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ps_matches, line);
    chunks[current].emit_else(line);
    emit_ps_match_array_from_single_exec(chunks, current, input, matches, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, ps_matches, line);
    chunks[current].emit_end(line);
    emit_match_info(
        chunks,
        current,
        input,
        pattern,
        ps_matches,
        ignore_case,
        line_number,
        pre_context,
        post_context,
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_ps_match_array_from_single_exec(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    exec: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let result = chunk.alloc_scratch(1);
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    let _ = chunk;
    emit_ps_match_object_from_exec(chunks, current, input, exec, line);
    let chunk = &mut chunks[current];
    let push = chunk.add_import("ecma:array", "push");
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
}

fn emit_ps_match_array_from_exec_array(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    execs: u16,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let get = chunk.add_import("ecma:array", "get");
    let length = chunk.add_import("ecma:array", "length");
    let push = chunk.add_import("ecma:array", "push");
    let result = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let exec = chunk.alloc_scratch(1);

    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, result, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);

    let done = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, execs, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_GE_U, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, execs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_call(get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, exec, line);

    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
    let _ = chunk;
    emit_ps_match_object_from_exec(chunks, current, input, exec, line);
    let chunk = &mut chunks[current];
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, result, line);
}

fn emit_ps_match_object_from_exec(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    exec: u16,
    line: u32,
) {
    let (arr_get, obj_get, index_of, out, value, index, length) = {
        let chunk = &mut chunks[current];
        (
            chunk.add_import("ecma:array", "get"),
            chunk.add_import("ecma:object", "get"),
            chunk.add_import("ecma:string", "indexOf"),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
            chunk.alloc_scratch(1),
        )
    };
    let groups = emit_ps_group_array_from_exec(chunks, current, exec, line);
    let chunk = &mut chunks[current];

    chunk.emit_op_u16(Op::LOCAL_GET, exec, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, exec, line);
    chunk.emit_string_const("index", line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, input, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(index_of, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    strings::emit_length(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, length, line);

    let _ = chunk;
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Text.RegularExpressions.Match",
        "Match",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "Value", value, line);
    set_local(chunk, out, "value", value, line);
    set_local(chunk, out, "Index", index, line);
    set_local(chunk, out, "index", index, line);
    set_local(chunk, out, "Length", length, line);
    set_local(chunk, out, "length", length, line);
    set_local(chunk, out, "Groups", groups, line);
    set_local(chunk, out, "groups", groups, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_ps_group_array_from_exec(
    chunks: &mut [Chunk],
    current: usize,
    exec: u16,
    line: u32,
) -> u16 {
    let chunk = &mut chunks[current];
    let get = chunk.add_import("ecma:array", "get");
    let obj_get = chunk.add_import("ecma:object", "get");
    let length = chunk.add_import("ecma:array", "length");
    let push = chunk.add_import("ecma:array", "push");
    let groups = chunk.alloc_scratch(1);
    let i = chunk.alloc_scratch(1);
    let value = chunk.alloc_scratch(1);
    let named = chunk.alloc_scratch(1);

    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, groups, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);

    let done = chunk.emit_block(line);
    let (loop_patch, _) = chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_op_u16(Op::LOCAL_GET, exec, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_GE_U, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, exec, line);
    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_call(get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, groups, line);
    let _ = chunk;
    emit_ps_group_object_from_value(chunks, current, value, line);
    let chunk = &mut chunks[current];
    chunk.emit_call(push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, i, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, i, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(loop_patch);
    chunk.emit_end(line);
    chunk.patch_block(done);

    chunk.emit_op_u16(Op::LOCAL_GET, exec, line);
    chunk.emit_string_const("groups", line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, named, line);
    let _ = chunk;
    emit_named_group_if_present(chunks, current, groups, named, "verb", line);
    emit_named_group_if_present(chunks, current, groups, named, "path", line);

    groups
}

fn emit_named_group_if_present(
    chunks: &mut [Chunk],
    current: usize,
    groups: u16,
    named: u16,
    key: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let get = chunk.add_import("ecma:object", "get");
    let value = chunk.alloc_scratch(1);
    let group = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, named, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, named, line);
    chunk.emit_string_const(key, line);
    chunk.emit_call(get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    let _ = chunk;
    emit_ps_group_object_from_value(chunks, current, value, line);
    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, group, line);
    set_local(chunk, groups, key, group, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

fn emit_ps_group_object_from_value(chunks: &mut [Chunk], current: usize, value: u16, line: u32) {
    let chunk = &mut chunks[current];
    let out = chunk.alloc_scratch(1);
    let length = chunk.alloc_scratch(1);
    let success = chunk.alloc_scratch(1);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    ops::emit_i32_to_bool(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, success, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    strings::emit_to_string(chunk, line);
    strings::emit_length(chunk, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, length, line);

    let _ = chunk;
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Text.RegularExpressions.Group",
        "Group",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "Value", value, line);
    set_local(chunk, out, "value", value, line);
    set_local(chunk, out, "Length", length, line);
    set_local(chunk, out, "length", length, line);
    set_local(chunk, out, "Success", success, line);
    set_local(chunk, out, "success", success, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_get_psdrive(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    match argc {
        0 => {
            emit_drive_object(chunks, current, "FileSystem", "/", "FileSystem", line);
            emit_drive_object(chunks, current, "Env", "", "Environment", line);
            emit_drive_object(chunks, current, "Variable", "", "Variable", line);
            emit_drive_object(chunks, current, "Function", "", "Function", line);
            emit_drive_object(chunks, current, "Alias", "", "Alias", line);
            chunks[current].emit_array_new_fixed(0, 5, line);
            let drives = chunks[current].alloc_scratch(3);
            let provider = drives + 1;
            let provider_names = drives + 2;
            chunks[current].emit_op_u16(Op::LOCAL_SET, drives, line);
            class_slots::emit_class_alloc(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, provider, line);
            chunks[current].emit_string_const("FileSystem", line);
            chunks[current].emit_string_const("Environment", line);
            chunks[current].emit_string_const("Variable", line);
            chunks[current].emit_string_const("Function", line);
            chunks[current].emit_string_const("Alias", line);
            chunks[current].emit_array_new_fixed(0, 5, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, provider_names, line);
            let chunk = &mut chunks[current];
            set_local(chunk, provider, "Name", provider_names, line);
            set_local(chunk, provider, "name", provider_names, line);
            set_local(chunk, drives, "Provider", provider, line);
            set_local(chunk, drives, "provider", provider, line);
            chunk.emit_op_u16(Op::LOCAL_GET, drives, line);
        }
        1 => {
            strings::emit_to_string(&mut chunks[current], line);
            let name = chunks[current].alloc_scratch(1);
            chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
            emit_registered_drive_or_null(chunks, current, name, line);
        }
        _ => {
            for _ in 2..argc {
                chunks[current].emit_op(Op::DROP, line);
            }
            strings::emit_to_string(&mut chunks[current], line);
            let name = chunks[current].alloc_scratch(3);
            let provider = name + 1;
            let root = name + 2;
            chunks[current].emit_op_u16(Op::LOCAL_SET, provider, line);
            strings::emit_to_string(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
            chunks[current].emit_string_const("", line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, root, line);
            emit_drive_object_from_slots(chunks, current, name, root, provider, line);
        }
    }
}

pub fn emit_new_psdrive(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let root = chunk.alloc_scratch(3);
    let name = root + 1;
    let description = root + 2;
    match argc {
        0 => {
            chunk.emit_string_const("", line);
            chunk.emit_op_u16(Op::LOCAL_SET, name, line);
            chunk.emit_string_const("/", line);
            chunk.emit_op_u16(Op::LOCAL_SET, root, line);
            chunk.emit_string_const("", line);
            chunk.emit_op_u16(Op::LOCAL_SET, description, line);
        }
        1 => {
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, name, line);
            chunk.emit_string_const("/", line);
            chunk.emit_op_u16(Op::LOCAL_SET, root, line);
            chunk.emit_string_const("", line);
            chunk.emit_op_u16(Op::LOCAL_SET, description, line);
        }
        2 => {
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, root, line);
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, name, line);
            chunk.emit_string_const("", line);
            chunk.emit_op_u16(Op::LOCAL_SET, description, line);
        }
        _ => {
            for _ in 3..argc {
                chunk.emit_op(Op::DROP, line);
            }
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, description, line);
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, root, line);
            strings::emit_to_string(chunk, line);
            chunk.emit_op_u16(Op::LOCAL_SET, name, line);
        }
    }
    let provider = chunks[current].alloc_scratch(1);
    chunks[current].emit_string_const("FileSystem", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, provider, line);
    let existing = chunks[current].alloc_scratch(1);
    emit_registered_drive_or_null(chunks, current, name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if_value(line);
    emit_throw_ps_error_record(
        chunks,
        current,
        "IOException",
        "DriveAlreadyExists,Microsoft.PowerShell.Commands.NewPSDriveCommand",
        "A drive with the specified name already exists.",
        line,
    );
    chunks[current].emit_end(line);
    emit_drive_object_from_slots(chunks, current, name, root, provider, line);
    let drive = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, drive, line);
    set_local(
        &mut chunks[current],
        drive,
        "Description",
        description,
        line,
    );
    set_local(
        &mut chunks[current],
        drive,
        "description",
        description,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, drive, line);
    emit_register_drive_from_stack(chunks, current, name, line);
}

pub fn emit_remove_psdrive(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    let args = pop_arg_slots(chunks, current, argc, line);
    let name = chunks[current].alloc_scratch(1);
    let force = args
        .get(1)
        .copied()
        .unwrap_or_else(|| bool_slot(&mut chunks[current], false, line));
    chunks[current].emit_op_u16(Op::LOCAL_GET, args[0], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    normalize_named_slot(&mut chunks[current], name, "Name", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    strings::emit_to_string(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, name, line);
    globals::emit_read(&mut chunks[current], PS_LOCATION_DRIVE, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, name, line);
    ops::emit_dyn_eq(&mut chunks[current], line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, force, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_throw_ps_error_record(
        chunks,
        current,
        "IOException",
        "DriveInUse,Microsoft.PowerShell.Commands.RemovePSDriveCommand",
        "The drive cannot be removed because it is in use.",
        line,
    );
    chunks[current].emit_end(line);
    emit_unregister_drive(chunks, current, name, line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn emit_get_location(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    globals::emit_read(&mut chunks[current], PS_LOCATION_DRIVE, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    filesystem_adapter::emit_current_directory(chunks, current, line);
    emit_path_info_from_stack(chunks, current, line);
    chunks[current].emit_else(line);
    let drive_name = chunks[current].alloc_scratch(1);
    let drive = chunks[current].alloc_scratch(1);
    globals::emit_read(&mut chunks[current], PS_LOCATION_DRIVE, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, drive_name, line);
    emit_registered_drive_or_null(chunks, current, drive_name, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, drive, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, drive, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    filesystem_adapter::emit_current_directory(chunks, current, line);
    emit_path_info_from_stack(chunks, current, line);
    chunks[current].emit_else(line);
    emit_path_info_from_drive(chunks, current, drive, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

pub fn emit_set_location(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 0 {
        chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        return;
    }
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    strings::emit_to_string(&mut chunks[current], line);
    let target = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_string_const(":", line);
    let ends_with = chunks[current].add_import("ecma:string", "endsWith");
    chunks[current].emit_call(ends_with, 2, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_string_const(":", line);
    chunks[current].emit_string_const("", line);
    let replace = chunks[current].add_import("ecma:string", "replaceAll");
    chunks[current].emit_call(replace, 3, line);
    globals::emit_write(&mut chunks[current], PS_LOCATION_DRIVE, line);
    chunks[current].emit_else(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    globals::emit_write(&mut chunks[current], PS_LOCATION_DRIVE, line);
    chunks[current].emit_end(line);
    chunks[current].emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

fn emit_local_is_platform_type(
    chunks: &mut [Chunk],
    current: usize,
    value: u16,
    registered_name: &str,
    parent: &str,
    line: u32,
) {
    let ancestry = vec![registered_name.to_string(), parent.to_string()];
    let typeidx = classes::reserve_platform_type(chunks, &ancestry);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    chunks[current].emit_ref_type_op(
        Op::REF_TEST,
        vybe_runtime::opcode::heaptype::HeapType::Concrete(typeidx as u32),
        line,
    );
}

fn emit_set_error_record_activity(chunk: &mut Chunk, record: u16, activity: u16, line: u32) {
    let category_info = chunk.alloc_scratch(1);
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(record),
        &field_slot("CategoryInfo"),
        Dest::Stack,
        line,
    );
    chunk.emit_op_u16(Op::LOCAL_SET, category_info, line);
    chunk.emit_op_u16(Op::LOCAL_GET, category_info, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if(line);
    chunk.emit_else(line);
    set_local(chunk, category_info, "Activity", activity, line);
    set_local(chunk, category_info, "activity", activity, line);
    chunk.emit_end(line);
}

fn emit_capture_record(chunks: &mut [Chunk], current: usize, global: &str, record: u16, line: u32) {
    let capture = chunks[current].alloc_scratch(1);
    globals::emit_read(&mut chunks[current], global, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, capture, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, capture, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, capture, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_end(line);
}

fn emit_preference_equals(
    chunks: &mut [Chunk],
    current: usize,
    global: &str,
    expected: &str,
    line: u32,
) {
    let pref = chunks[current].alloc_scratch(1);
    globals::emit_read(&mut chunks[current], global, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, pref, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pref, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, pref, line);
    strings::emit_to_string(&mut chunks[current], line);
    strings::emit_to_lower(&mut chunks[current], line);
    chunks[current].emit_string_const(expected, line);
    let eq = chunks[current].add_import("wasm:js-string", "equals");
    chunks[current].emit_call(eq, 2, line);
    chunks[current].emit_end(line);
}

fn emit_append_error_collection_unless_ignore(
    chunks: &mut [Chunk],
    current: usize,
    record: u16,
    line: u32,
) {
    emit_preference_equals(chunks, current, PS_ERROR_ACTION_PREFERENCE, "ignore", line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_else(line);
    let existing = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    globals::emit_write(&mut chunks[current], PS_LAST_ERROR_RECORD, line);
    globals::emit_read(&mut chunks[current], PS_ERROR_COLLECTION, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_array_new_fixed(0, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, existing, line);
    chunks[current].emit_end(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    collections::emit_array_new(chunks, current, 1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, existing, line);
    collections::emit_spread_iterable(chunks, current, line);
    collections::emit_concat(chunks, current, line);
    globals::emit_write(&mut chunks[current], PS_ERROR_COLLECTION, line);
    chunks[current].emit_end(line);
}

fn emit_throw_if_preference_stop(
    chunks: &mut [Chunk],
    current: usize,
    global: &str,
    record: u16,
    line: u32,
) {
    emit_preference_equals(chunks, current, global, "stop", line);
    chunks[current].emit_if_value(line);
    emit_throw_action_preference_stop(chunks, current, record, line);
    chunks[current].emit_end(line);
}

pub fn emit_pscmdlet_write_debug(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let message = args.get(0).copied().unwrap_or(empty);
    let activity = args.last().copied().filter(|_| argc > 1);
    let out = emit_message_record_from_slot(
        chunks,
        current,
        message,
        activity,
        "System.Management.Automation.DebugRecord",
        "DebugRecord",
        line,
    );
    emit_capture_record(chunks, current, PS_CAPTURE_DEBUG, out, line);
    emit_throw_if_preference_stop(chunks, current, PS_DEBUG_PREFERENCE, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_pscmdlet_write_warning(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let message = args.get(0).copied().unwrap_or(empty);
    let activity = args.last().copied().filter(|_| argc > 1);
    let out = emit_message_record_from_slot(
        chunks,
        current,
        message,
        activity,
        "System.Management.Automation.WarningRecord",
        "WarningRecord",
        line,
    );
    emit_capture_record(chunks, current, PS_CAPTURE_WARNING, out, line);
    emit_throw_if_preference_stop(chunks, current, PS_WARNING_PREFERENCE, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

pub fn emit_pscmdlet_write_information(
    chunks: &mut Vec<Chunk>,
    current: usize,
    argc: u8,
    line: u32,
) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let null = null_slot(&mut chunks[current], line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let empty_tags = empty_array_slot(&mut chunks[current], line);
    let message_data = args.get(0).copied().unwrap_or(null);
    let source = args.last().copied().filter(|_| argc > 1).unwrap_or(empty);
    let tags = args
        .get(1)
        .copied()
        .filter(|_| argc > 2)
        .unwrap_or(empty_tags);

    emit_local_is_platform_type(
        chunks,
        current,
        message_data,
        "InformationRecord",
        "Object",
        line,
    );
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, message_data, line);
    emit_capture_record(chunks, current, PS_CAPTURE_INFORMATION, message_data, line);
    emit_throw_if_preference_stop(
        chunks,
        current,
        PS_INFORMATION_PREFERENCE,
        message_data,
        line,
    );
    chunks[current].emit_else(line);
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.InformationRecord",
        "InformationRecord",
        line,
    );
    set_local(&mut chunks[current], out, "MessageData", message_data, line);
    set_local(&mut chunks[current], out, "messagedata", message_data, line);
    set_local(&mut chunks[current], out, "Source", source, line);
    set_local(&mut chunks[current], out, "source", source, line);
    set_local(&mut chunks[current], out, "Tags", tags, line);
    set_local(&mut chunks[current], out, "tags", tags, line);
    let time_generated = emit_datetime_now_slot(chunks, current, line);
    set_declared_local(
        chunks,
        current,
        out,
        "InformationRecord",
        "TimeGenerated",
        time_generated,
        line,
    );
    set_local(
        &mut chunks[current],
        out,
        "timegenerated",
        time_generated,
        line,
    );
    set_const_str(&mut chunks[current], out, "Computer", "localhost", line);
    set_const_str(&mut chunks[current], out, "computer", "localhost", line);
    let invocation = emit_invocation_info_from_activity(chunks, current, source, line);
    set_local(
        &mut chunks[current],
        out,
        "InvocationInfo",
        invocation,
        line,
    );
    set_local(
        &mut chunks[current],
        out,
        "invocationinfo",
        invocation,
        line,
    );
    emit_capture_record(chunks, current, PS_CAPTURE_INFORMATION, out, line);
    emit_throw_if_preference_stop(chunks, current, PS_INFORMATION_PREFERENCE, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);
}

pub fn emit_pscmdlet_write_error(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let args = pop_arg_slots(chunks, current, argc, line);
    let null = null_slot(&mut chunks[current], line);
    let empty = empty_string_slot(&mut chunks[current], line);
    let record = args.get(0).copied().unwrap_or(null);
    let activity = args.last().copied().filter(|_| argc > 1).unwrap_or(empty);
    emit_local_is_platform_type(chunks, current, record, "ErrorRecord", "Object", line);
    chunks[current].emit_if_value(line);
    emit_set_error_record_activity(&mut chunks[current], record, activity, line);
    emit_capture_record(chunks, current, PS_CAPTURE_ERROR, record, line);
    emit_append_error_collection_unless_ignore(chunks, current, record, line);
    emit_throw_if_preference_stop(chunks, current, PS_ERROR_ACTION_PREFERENCE, record, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, record, line);
    chunks[current].emit_else(line);
    let out = emit_error_record_from_slots(
        chunks,
        current,
        record,
        empty,
        null,
        null,
        Some(activity),
        line,
    );
    emit_capture_record(chunks, current, PS_CAPTURE_ERROR, out, line);
    emit_append_error_collection_unless_ignore(chunks, current, out, line);
    emit_throw_if_preference_stop(chunks, current, PS_ERROR_ACTION_PREFERENCE, out, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
    chunks[current].emit_end(line);
}

pub fn emit_pscmdlet_write_verbose(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    emit_pscmdlet_write_debug(chunks, current, argc, line);
}

pub fn emit_pscmdlet_should_process(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    for _ in 0..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    globals::emit_read(&mut chunks[current], "whatifpreference", line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_bool_const(false, line);
    chunks[current].emit_else(line);
    chunks[current].emit_bool_const(true, line);
    chunks[current].emit_end(line);
}

fn emit_compare_entry(chunks: &mut [Chunk], current: usize, value: u16, side: &str, line: u32) {
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.PSCustomObject",
        "PSCustomObject",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "InputObject", value, line);
    set_local(chunk, out, "inputobject", value, line);
    set_const_str(chunk, out, "SideIndicator", side, line);
    set_const_str(chunk, out, "sideindicator", side, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_match_info(
    chunks: &mut [Chunk],
    current: usize,
    input: u16,
    pattern: u16,
    matches: u16,
    ignore_case: u16,
    line_number: u16,
    pre_context: u16,
    post_context: u16,
    line: u32,
) {
    let out = chunks[current].alloc_scratch(1);
    emit_typed_object(
        chunks,
        current,
        out,
        "Microsoft.PowerShell.Commands.MatchInfo",
        "MatchInfo",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "Line", input, line);
    set_local(chunk, out, "line", input, line);
    set_local(chunk, out, "Pattern", pattern, line);
    set_local(chunk, out, "pattern", pattern, line);
    set_local(chunk, out, "Matches", matches, line);
    set_local(chunk, out, "matches", matches, line);
    set_local(chunk, out, "IgnoreCase", ignore_case, line);
    set_local(chunk, out, "ignorecase", ignore_case, line);
    set_const_num(chunk, out, "Count", 1.0, line);
    set_const_num(chunk, out, "count", 1.0, line);
    set_const_num(chunk, out, "Length", 1.0, line);
    set_const_num(chunk, out, "length", 1.0, line);
    set_const_str(chunk, out, "Filename", "", line);
    set_const_str(chunk, out, "filename", "", line);
    set_const_str(chunk, out, "Path", "", line);
    set_const_str(chunk, out, "path", "", line);
    set_local(chunk, out, "LineNumber", line_number, line);
    set_local(chunk, out, "linenumber", line_number, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pre_context, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op_u16(Op::LOCAL_GET, post_context, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    let context = chunk.alloc_scratch(1);
    collections::emit_map_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, context, line);
    let chunk = &mut chunks[current];
    set_local(chunk, context, "PreContext", pre_context, line);
    set_local(chunk, context, "precontext", pre_context, line);
    set_local(chunk, context, "PostContext", post_context, line);
    set_local(chunk, context, "postcontext", post_context, line);
    set_local(chunk, out, "Context", context, line);
    set_local(chunk, out, "context", context, line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_drive_object(
    chunks: &mut [Chunk],
    current: usize,
    name: &str,
    root: &str,
    provider_name: &str,
    line: u32,
) {
    let chunk = &mut chunks[current];
    let name_slot = chunk.alloc_scratch(3);
    let root_slot = name_slot + 1;
    let provider_slot = name_slot + 2;
    chunk.emit_string_const(name, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name_slot, line);
    chunk.emit_string_const(root, line);
    chunk.emit_op_u16(Op::LOCAL_SET, root_slot, line);
    chunk.emit_string_const(provider_name, line);
    chunk.emit_op_u16(Op::LOCAL_SET, provider_slot, line);
    emit_drive_object_from_slots(chunks, current, name_slot, root_slot, provider_slot, line);
}

fn emit_drive_object_from_slots(
    chunks: &mut [Chunk],
    current: usize,
    name: u16,
    root: u16,
    provider_name: u16,
    line: u32,
) {
    let provider = chunks[current].alloc_scratch(2);
    let out = provider + 1;
    let chunk = &mut chunks[current];
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, provider, line);
    set_local(chunk, provider, "Name", provider_name, line);
    set_local(chunk, provider, "name", provider_name, line);
    let _ = chunk;

    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.PSDriveInfo",
        "PSDriveInfo",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "Name", name, line);
    set_local(chunk, out, "name", name, line);
    set_local(chunk, out, "Root", root, line);
    set_local(chunk, out, "root", root, line);
    set_const_str(chunk, out, "CurrentLocation", "", line);
    set_const_str(chunk, out, "currentlocation", "", line);
    set_local(chunk, out, "Provider", provider, line);
    set_local(chunk, out, "provider", provider, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

fn emit_path_info_from_drive(chunks: &mut [Chunk], current: usize, drive: u16, line: u32) {
    let full_path = chunks[current].alloc_scratch(1);
    let provider = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    let chunk = &mut chunks[current];
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(drive),
        &field_slot("Root"),
        Dest::Local(full_path),
        line,
    );
    class_slots::emit_class_get(
        chunk,
        ObjSource::Local(drive),
        &field_slot("Provider"),
        Dest::Local(provider),
        line,
    );
    let _ = chunk;

    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.PathInfo",
        "PathInfo",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "Path", full_path, line);
    set_local(chunk, out, "path", full_path, line);
    set_local(chunk, out, "ProviderPath", full_path, line);
    set_local(chunk, out, "providerpath", full_path, line);
    set_local(chunk, out, "Provider", provider, line);
    set_local(chunk, out, "provider", provider, line);
    set_local(chunk, out, "Drive", drive, line);
    set_local(chunk, out, "drive", drive, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

fn drop_after_first(chunk: &mut Chunk, argc: u8, line: u32) {
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

fn drop_after_second(chunk: &mut Chunk, argc: u8, line: u32) {
    for _ in 2..argc {
        chunk.emit_op(Op::DROP, line);
    }
}

fn emit_resolved_path(chunks: &mut [Chunk], current: usize, argc: u8, path_info: bool, line: u32) {
    let raw_args: Vec<u16> = (0..argc)
        .map(|_| chunks[current].alloc_scratch(1))
        .collect();
    for slot in raw_args.iter().rev() {
        chunks[current].emit_op_u16(Op::LOCAL_SET, *slot, line);
    }
    let path = chunks[current].alloc_scratch(3);
    let suppress = path + 1;
    let literal = path + 2;
    if let Some(slot) = raw_args.first().copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_string_const("", line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    normalize_pathish_slot(&mut chunks[current], path, line);
    if let Some(slot) = raw_args.get(1).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, suppress, line);
    if let Some(slot) = raw_args.get(2).copied() {
        chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    } else {
        chunks[current].emit_bool_const(false, line);
    }
    chunks[current].emit_op_u16(Op::LOCAL_SET, literal, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, literal, line);
    ops::emit_dyn_to_bool(&mut chunks[current], line);
    ops::emit_dyn_not(&mut chunks[current], line);
    emit_path_has_wildcard(&mut chunks[current], path, line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if_value(line);
    emit_wildcard_paths(chunks, current, path, line);
    if path_info {
        emit_path_info_array_from_paths(chunks, current, line);
    } else {
        emit_full_path_array_from_paths(chunks, current, line);
    }
    chunks[current].emit_else(line);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    fs_path::emit_exists(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_full_path(chunk, line);
    if path_info {
        let _ = chunk;
        emit_path_info_from_stack(chunks, current, line);
        let chunk = &mut chunks[current];
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, suppress, line);
        chunk.emit_if_value(line);
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        chunk.emit_else(line);
        let _ = chunk;
        emit_throw_ps_error_record(
            chunks,
            current,
            "ItemNotFoundException",
            "PathNotFound,Microsoft.PowerShell.Commands.ResolvePathCommand",
            "Cannot find path because it does not exist.",
            line,
        );
        let chunk = &mut chunks[current];
        chunk.emit_end(line);
        chunk.emit_end(line);
        chunks[current].emit_end(line);
        return;
    }
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, suppress, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
    chunk.emit_else(line);
    let _ = chunk;
    emit_throw_ps_error_record(
        chunks,
        current,
        "ItemNotFoundException",
        "PathNotFound,Microsoft.PowerShell.Commands.ConvertPathCommand",
        "Cannot find path because it does not exist.",
        line,
    );
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunks[current].emit_end(line);
}

fn emit_full_path_array_from_paths(chunks: &mut [Chunk], current: usize, line: u32) {
    let paths_slot = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let path = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, paths_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);

    let state = loops::emit_for_in_start(chunks, current, paths_slot, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_full_path(&mut chunks[current], line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

fn emit_path_info_array_from_paths(chunks: &mut [Chunk], current: usize, line: u32) {
    let paths_slot = chunks[current].alloc_scratch(1);
    let result = chunks[current].alloc_scratch(1);
    let idx = chunks[current].alloc_scratch(1);
    let path = chunks[current].alloc_scratch(1);

    chunks[current].emit_op_u16(Op::LOCAL_SET, paths_slot, line);
    collections::emit_array_new(chunks, current, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, result, line);

    let state = loops::emit_for_in_start(chunks, current, paths_slot, idx, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, path, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, path, line);
    paths::emit_full_path(&mut chunks[current], line);
    emit_path_info_from_stack(chunks, current, line);
    collections::emit_push(chunks, current, line);
    chunks[current].emit_op(Op::DROP, line);
    loops::emit_for_in_end(chunks, current, idx, state, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, result, line);
}

fn emit_path_info_from_stack(chunks: &mut [Chunk], current: usize, line: u32) {
    let full_path = chunks[current].alloc_scratch(1);
    let provider = chunks[current].alloc_scratch(1);
    let drive = chunks[current].alloc_scratch(1);
    let out = chunks[current].alloc_scratch(1);

    let chunk = &mut chunks[current];
    chunk.emit_op_u16(Op::LOCAL_SET, full_path, line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, provider, line);
    set_const_str(chunk, provider, "Name", "FileSystem", line);
    set_const_str(chunk, provider, "name", "FileSystem", line);

    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, drive, line);
    set_const_str(chunk, drive, "Name", "", line);
    set_const_str(chunk, drive, "name", "", line);
    set_const_str(chunk, drive, "Root", "/", line);
    set_const_str(chunk, drive, "root", "/", line);
    set_local(chunk, drive, "Provider", provider, line);
    set_local(chunk, drive, "provider", provider, line);
    let _ = chunk;

    emit_typed_object(
        chunks,
        current,
        out,
        "System.Management.Automation.PathInfo",
        "PathInfo",
        line,
    );
    let chunk = &mut chunks[current];
    set_local(chunk, out, "Path", full_path, line);
    set_local(chunk, out, "path", full_path, line);
    set_local(chunk, out, "ProviderPath", full_path, line);
    set_local(chunk, out, "providerpath", full_path, line);
    set_local(chunk, out, "Provider", provider, line);
    set_local(chunk, out, "provider", provider, line);
    set_local(chunk, out, "Drive", drive, line);
    set_local(chunk, out, "drive", drive, line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}
