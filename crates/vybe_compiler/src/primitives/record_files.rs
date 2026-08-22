//! Record files — `FileDecl` and `RecordTransfer` lowered onto WASI 0.3.1.
//!
//! One emitter for every language that has record I/O. COBOL `SELECT`/`FD` +
//! `READ`/`WRITE`/`REWRITE`/`DELETE`, VB `Open`/`Get`/`Put`, Pascal
//! `Reset`/`Read`/`Write`, Fortran `OPEN`/`READ(rec=)` are the same two
//! operations over different addressing modes, which is why they used to be
//! nine AST nodes that each knew a little and none knew the record layout.
//!
//! ## Where the width comes from
//!
//! `FileDecl.record` names a `StructDecl` whose fields carry
//! [`vybe_ast::FieldStorage`]. The width of a record is the sum of its
//! members' byte extents, read from the DECLARATION at emit time — never
//! re-derived from a PICTURE, a `String * n`, or a `character(len=)` here.
//! A language states the extent once, in its walker; this reads it.
//!
//! ## The transport
//!
//! `wasi:filesystem/types` and nothing else:
//!
//! - `[method]descriptor.open-at(parent, path-flags, path, open-flags, desc-flags)`
//! - `[method]descriptor.write-via-stream(data, offset)`
//! - `[method]descriptor.read-via-stream(offset)`
//!
//! WASI 0.3.1 deleted positioned `descriptor.read`/`descriptor.write` — the
//! via-stream pair is the only door, and the WIT says they behave like
//! `pread`/`pwrite`. That is what makes record *n* at *n × width* expressible
//! without a seek: the offset is an argument.

use crate::primitives::Compiler;
use vybe_ast::{
    Expression, FieldStorage, FileAccess, FileOrganization, Justify, OpenMode, RecordAddress,
    RecordDirection, RecordKey, TypeRef,
};
use vybe_runtime::opcode::Op;

// ── The handle ──────────────────────────────────────────────────────────
//
// A declared file binds its own NAME to this map. Not a number: a file's
// identity is its declaration, and only VB6 identifies one by an integer the
// programmer invents.

/// The open `descriptor` resource.
const F_DESC: &str = "descriptor";
/// Bytes per record, fixed at declaration.
const F_WIDTH: &str = "width";
/// Next record index for sequential access, 0-based. Advanced by every
/// `RecordAddress::Next` transfer.
const F_POS: &str = "position";
/// The path, retained so a diagnostic can name the file rather than a handle.
const F_PATH: &str = "path";

// WASI `open-flags` / `descriptor-flags`, `wasi:filesystem/types`.
const OPEN_CREATE: i32 = 1;
const OPEN_TRUNCATE: i32 = 8;
const DESC_READ: i32 = 1;
const DESC_WRITE: i32 = 2;

/// What opening in this mode asks of the host.
///
/// COBOL's four `OPEN` verbs, VB's `For`, and Pascal's
/// `Reset`/`Rewrite`/`Append` all land here, and the mapping is the whole of
/// what distinguishes them: `OUTPUT` truncating is why writing over a full
/// file leaves an empty one, and `I-O` not truncating is why `REWRITE` has
/// something to rewrite.
fn open_flags_for(mode: OpenMode) -> (i32, i32) {
    match mode {
        OpenMode::Read => (0, DESC_READ),
        OpenMode::Write => (OPEN_CREATE | OPEN_TRUNCATE, DESC_WRITE),
        OpenMode::ReadWrite => (OPEN_CREATE, DESC_READ | DESC_WRITE),
        OpenMode::Append => (OPEN_CREATE, DESC_WRITE),
    }
}

impl Compiler {
    /// The declared byte extent of a record type: the sum of its members'.
    ///
    /// `None` when the type is unknown or any member states no extent — a
    /// record with one unmeasurable field has no width at all, and answering
    /// with the sum of the rest would put every later record at the wrong
    /// offset with no symptom but wrong data.
    fn record_width(&self, record: &TypeRef) -> Option<u32> {
        let name = record_type_name(record)?;
        let class = self
            .normalized_classes
            .get(&name)
            .or_else(|| self.normalized_classes.get(&self.canon(&name)))?;
        let mut total = 0u32;
        for field in &class.instance_fields {
            total = total.checked_add(field.storage?.bytes)?;
        }
        (total > 0).then_some(total)
    }

    /// The members of a record type, in declaration order, with their extents.
    fn record_layout(&self, record: &TypeRef) -> Option<Vec<(String, FieldStorage)>> {
        let name = record_type_name(record)?;
        let class = self
            .normalized_classes
            .get(&name)
            .or_else(|| self.normalized_classes.get(&self.canon(&name)))?;
        class
            .instance_fields
            .iter()
            .map(|f| Some((f.name.clone(), f.storage?)))
            .collect()
    }

    /// `FileDecl` — open the file and bind its name to a handle.
    pub(super) fn compile_file_decl(
        &mut self,
        name: &str,
        path: &Expression,
        record: &TypeRef,
        organization: FileOrganization,
        _access: FileAccess,
        mode: OpenMode,
        keys: &[RecordKey],
    ) -> Result<(), String> {
        if !keys.is_empty() {
            return Err(format!(
                "FileDecl `{name}`: INDEXED files are not lowered yet — a key lookup is a scan \
                 or a side index, not an offset, so it needs its own emitter rather than the \
                 positioned transfer this module does"
            ));
        }

        // LINE organization has no fixed width, so a record type is optional
        // there and required everywhere else. Demanding one uniformly would
        // reject `ORGANIZATION LINE SEQUENTIAL`, which is most COBOL text I/O.
        let width = match organization {
            FileOrganization::Line => 0,
            _ => self.record_width(record).ok_or_else(|| {
                format!(
                    "FileDecl `{name}`: record type `{}` has no declared byte extent — every \
                     member needs `FieldStorage` before record {} × width can be addressed",
                    record_type_name(record).unwrap_or_else(|| "?".into()),
                    'n'
                )
            })?,
        };

        let line = self.line;
        let (open_flags, desc_flags) = open_flags_for(mode);

        let path_slot = self.chunk().alloc_scratch(1);
        self.compile_expr(path)?;
        self.chunk().emit_op_u16(Op::LOCAL_SET, path_slot, line);

        // The preopen directory. WASI resolves every path against a parent
        // descriptor — there is no absolute `open`, by design — and vybe
        // preopens the cwd as ".".
        let preopens = self
            .chunk()
            .add_import("wasi:filesystem/preopens", "get-directories");
        let at = self.chunk().add_import("ecma:array", "at");
        let open_at = self
            .chunk()
            .add_import("wasi:filesystem/types", "[method]descriptor.open-at");

        self.chunk().emit_call(preopens, 0, line);
        self.chunk().emit_i32_const(0, line);
        self.chunk().emit_call(at, 2, line); // first tuple<descriptor, string>
        self.chunk().emit_i32_const(0, line);
        self.chunk().emit_call(at, 2, line); // its descriptor

        self.chunk().emit_i32_const(0, line); // path-flags: no symlink follow
        self.chunk().emit_op_u16(Op::LOCAL_GET, path_slot, line);
        self.chunk().emit_i32_const(open_flags, line);
        self.chunk().emit_i32_const(desc_flags, line);
        self.chunk().emit_call(open_at, 5, line);

        let desc_slot = self.chunk().alloc_scratch(1);
        self.chunk().emit_op_u16(Op::LOCAL_SET, desc_slot, line);

        let current = self.current;
        let handle_slot = self.chunk().alloc_scratch(1);
        super::collections::emit_map_new(&mut self.chunks, current, line);
        self.chunk().emit_op_u16(Op::LOCAL_SET, handle_slot, line);

        // APPEND starts at the end, and "the end" in records is
        // `size / width` — so this is the one mode that has to ask the file
        // how big it already is. Every other mode starts at record 0.
        //
        // Done here rather than at the first write because the position is a
        // property of the OPEN: COBOL `EXTEND` fixes where writing resumes,
        // and re-deriving it per transfer would let an intervening write by
        // anyone else move it.
        let start_pos = if mode == OpenMode::Append && width > 0 {
            let stat = self
                .chunk()
                .add_import("wasi:filesystem/types", "[method]descriptor.stat");
            let pos_slot = self.chunk().alloc_scratch(1);
            self.chunk().emit_op_u16(Op::LOCAL_GET, desc_slot, line);
            self.chunk().emit_call(stat, 1, line);
            self.chunk().emit_string_const("size", line);
            super::collections::emit_get(&mut self.chunks, current, line);
            self.chunk().emit_f64_const(width as f64, line);
            self.chunk().emit_op(Op::F64_DIV, line);
            super::math::emit_floor(&mut self.chunks[current], line);
            self.chunk().emit_op_u16(Op::LOCAL_SET, pos_slot, line);
            HandleInit::Slot(pos_slot)
        } else {
            HandleInit::Int(0)
        };

        for (field, emit) in [
            (F_DESC, HandleInit::Slot(desc_slot)),
            (F_WIDTH, HandleInit::Int(width as i64)),
            (F_POS, start_pos),
            (F_PATH, HandleInit::Slot(path_slot)),
        ] {
            self.chunk().emit_op_u16(Op::LOCAL_GET, handle_slot, line);
            self.chunk().emit_string_const(field, line);
            match emit {
                HandleInit::Slot(slot) => self.chunk().emit_op_u16(Op::LOCAL_GET, slot, line),
                // f64, not i32: `position` and `width` are multiplied and
                // incremented below, and mixing an i32 handle field into
                // `F64_MUL` is the kind of thing that works until it doesn't.
                HandleInit::Int(v) => self.chunk().emit_f64_const(v as f64, line),
            }
            super::collections::emit_set(&mut self.chunks, current, line);
            self.chunk().emit_op(Op::DROP, line);
        }

        self.chunk().emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        self.compile_assign_target(&Expression::ident(name))
    }

    /// `RecordTransfer` — one positioned read or write.
    pub(super) fn compile_record_transfer(
        &mut self,
        file: &Expression,
        record_type: &TypeRef,
        direction: RecordDirection,
        at: &RecordAddress,
        record: Option<&Expression>,
        status: Option<&Expression>,
    ) -> Result<(), String> {
        // A key address is a LOOKUP, not an offset: finding it means scanning
        // the file or consulting an index. Folding it in here would make the
        // positioned path pretend to answer a question it cannot.
        if let RecordAddress::Key { .. } = at {
            return Err(
                "RecordTransfer: keyed addressing needs an index or a scan, which the positioned \
                 emitter cannot express — INDEXED files are not lowered yet"
                    .to_string(),
            );
        }

        match direction {
            RecordDirection::Write | RecordDirection::Rewrite => {
                self.emit_record_write(file, record_type, at, record)?;
                // A positioned write past the end extends the file, so there
                // is no failure this path can report — the only status it can
                // honestly claim is success.
                self.emit_status_store(status, "00")
            }
            RecordDirection::Read => self.emit_record_read(file, record_type, at, record, status),
            RecordDirection::Delete => Err(
                "RecordTransfer DELETE: removing a record means rewriting the file or marking \
                 the slot, neither of which is a positioned write — not lowered yet"
                    .to_string(),
            ),
        }
    }

    /// A positioned read: `width` bytes at `index × width`, split back into the
    /// record's fields by the layout the type declared.
    ///
    /// `read-via-stream` answers `tuple<stream<u8>, future<result<_,
    /// error-code>>>` and element 0 is the readable end — the same shape every
    /// 0.3.1 stream producer uses. The future carries how the read ENDED; it
    /// is dropped here because a record read has nowhere to report an
    /// error-code yet, and inventing a channel for it would be a second file
    /// status model beside the one the languages already have.
    fn emit_record_read(
        &mut self,
        file: &Expression,
        record_type: &TypeRef,
        at: &RecordAddress,
        record: Option<&Expression>,
        status: Option<&Expression>,
    ) -> Result<(), String> {
        let Some(target) = record else {
            return Err("RecordTransfer READ: no record to read into".to_string());
        };
        let layout = self.record_layout(record_type).ok_or_else(|| {
            "RecordTransfer READ: the file's record type has no declared layout, so there is no \
             way to say where one field ends and the next begins"
                .to_string()
        })?;

        let line = self.line;
        let current = self.current;

        let handle_slot = self.chunk().alloc_scratch(1);
        self.compile_expr(file)?;
        self.chunk().emit_op_u16(Op::LOCAL_SET, handle_slot, line);

        let desc_slot = self.chunk().alloc_scratch(1);
        self.emit_handle_field(handle_slot, F_DESC);
        self.chunk().emit_op_u16(Op::LOCAL_SET, desc_slot, line);

        let width_slot = self.chunk().alloc_scratch(1);
        self.emit_handle_field(handle_slot, F_WIDTH);
        self.chunk().emit_op_u16(Op::LOCAL_SET, width_slot, line);

        let offset_slot = self.chunk().alloc_scratch(1);
        self.emit_record_offset(handle_slot, width_slot, offset_slot, at)?;

        let read_via = self
            .chunk()
            .add_import("wasi:filesystem/types", "[method]descriptor.read-via-stream");
        let at_idx = self.chunk().add_import("ecma:array", "at");
        let end_slot = self.chunk().alloc_scratch(1);
        let text_slot = self.chunk().alloc_scratch(1);

        self.chunk().emit_op_u16(Op::LOCAL_GET, desc_slot, line);
        self.chunk().emit_op_u16(Op::LOCAL_GET, offset_slot, line);
        self.chunk().emit_call(read_via, 2, line);
        self.chunk().emit_i32_const(0, line);
        self.chunk().emit_call(at_idx, 2, line);
        self.chunk().emit_op_u16(Op::LOCAL_SET, end_slot, line);

        // A readable end is an i32 HANDLE. Anything else means the transfer
        // never started — `open-at` on a file that isn't there answers an
        // `error-code`, and `read-via-stream` on that answers another one
        // rather than a tuple, so element 0 is not a stream.
        //
        // Tested rather than assumed because `canon stream.read` TRAPS on a
        // non-handle ("handle is not a readable stream end"), which turns a
        // missing input file — an ordinary condition COBOL reports as a file
        // status — into a crash that takes the program with it.
        let is_number = self.chunk().add_import("wasm:js-number", "test");
        self.chunk().emit_op_u16(Op::LOCAL_GET, end_slot, line);
        self.chunk().emit_call(is_number, 1, line);
        self.chunk().emit_if(line);
        self.chunk().emit_op_u16(Op::LOCAL_GET, end_slot, line);
        super::io::emit_read_stream_to_string(&mut self.chunks[current], line);
        self.chunk().emit_op_u16(Op::LOCAL_SET, text_slot, line);
        self.chunk().emit_else(line);
        // No stream, so no record. That reads as end-of-file below, which is
        // what every caller of this already knows how to handle.
        self.chunk().emit_string_const("", line);
        self.chunk().emit_op_u16(Op::LOCAL_SET, text_slot, line);
        self.chunk().emit_end(line);

        // Field n starts where field n-1 ended. The cursor is a COMPILE-TIME
        // running total, not a runtime one, because every extent is declared:
        // this is the whole payoff of carrying `FieldStorage` through
        // normalization rather than asking the language at each site.
        let record_slot = self.chunk().alloc_scratch(1);
        self.compile_expr(target)?;
        self.chunk().emit_op_u16(Op::LOCAL_SET, record_slot, line);

        // AT END, tested BEFORE anything is written into the record.
        //
        // A read at or past the end drains an empty stream, and COBOL leaves
        // the record area UNCHANGED at end-of-file — it does not blank it.
        // Splitting an empty string into fields first would write `""` into
        // every one and convert the numerics, which is both wrong and a trap:
        // `Number("")` is not an integer and the conversion faults rather than
        // producing a wrong answer.
        let str_len = self.chunk().add_import("ecma:string", "length");
        self.chunk().emit_op_u16(Op::LOCAL_GET, text_slot, line);
        self.chunk().emit_call(str_len, 1, line);
        self.chunk().emit_i32_const(0, line);
        self.chunk().emit_op(Op::I32_EQ, line);
        self.chunk().emit_if(line);
        self.emit_status_store(status, "10")?;
        self.chunk().emit_else(line);

        // Field n starts where field n-1 ended. The cursor is a COMPILE-TIME
        // running total, not a runtime one, because every extent is declared:
        // this is the whole payoff of carrying `FieldStorage` through
        // normalization rather than asking the language at each site.
        let mut cursor = 0u32;
        for (name, storage) in &layout {
            let end = cursor + storage.bytes;
            let key = self.canon(name);
            self.chunk().emit_op_u16(Op::LOCAL_GET, record_slot, line);
            self.chunk().emit_string_const(&key, line);
            self.chunk().emit_op_u16(Op::LOCAL_GET, text_slot, line);
            self.chunk().emit_i32_const(cursor as i32, line);
            self.chunk().emit_i32_const(end as i32, line);
            super::strings::emit_substring(&mut self.chunks[current], line);
            self.emit_from_storage(*storage);
            super::collections::emit_set(&mut self.chunks, current, line);
            self.chunk().emit_op(Op::DROP, line);
            cursor = end;
        }
        self.emit_status_store(status, "00")?;
        self.chunk().emit_end(line);

        Ok(())
    }

    /// Store a two-character file status, when the program asked for one.
    fn emit_status_store(
        &mut self,
        status: Option<&Expression>,
        code: &str,
    ) -> Result<(), String> {
        let Some(status) = status else {
            return Ok(());
        };
        let line = self.line;
        self.chunk().emit_string_const(code, line);
        let target = status.clone();
        self.compile_assign_target(&target)
    }

    /// One field's stored characters → the value it denotes.
    /// Stack: `[string]` → `[value]`.
    ///
    /// The inverse of [`Self::emit_to_storage`], and only for numerics: text
    /// fields ARE their characters, padding included, because a COBOL
    /// `PIC X(10)` holding `"AB"` is eight trailing spaces and every reader
    /// downstream expects them.
    fn emit_from_storage(&mut self, storage: FieldStorage) {
        if storage.justify != Justify::Right {
            return;
        }
        let line = self.line;
        let current = self.current;
        super::convert::emit_to_float(&mut self.chunks[current], line);
        if storage.decimal_places > 0 {
            let scale = 10f64.powi(storage.decimal_places as i32);
            self.chunk().emit_f64_const(scale, line);
            self.chunk().emit_op(Op::F64_DIV, line);
        }
    }

    /// `offset_slot` = where this record's bytes begin.
    ///
    /// `write-via-stream` says writing past the end is legal and zero-fills
    /// the gap, which is what makes a sparse relative file work without a
    /// separate extend — so this is the only positioning either direction
    /// needs.
    fn emit_record_offset(
        &mut self,
        handle_slot: u16,
        width_slot: u16,
        offset_slot: u16,
        at: &RecordAddress,
    ) -> Result<(), String> {
        let line = self.line;
        let current = self.current;
        match at {
            RecordAddress::Current => {
                // One record BEHIND the position: a read left it on the
                // record AFTER the one just handed over, and that one is what
                // a rewrite replaces. The position is NOT advanced — a
                // rewrite replaces a record, it does not consume one.
                self.emit_handle_field(handle_slot, F_POS);
                self.chunk().emit_f64_const(1.0, line);
                self.chunk().emit_op(Op::F64_SUB, line);
                self.chunk().emit_op_u16(Op::LOCAL_GET, width_slot, line);
                self.chunk().emit_op(Op::F64_MUL, line);
                self.chunk().emit_op_u16(Op::LOCAL_SET, offset_slot, line);
            }
            RecordAddress::Next => {
                self.emit_handle_field(handle_slot, F_POS);
                self.chunk().emit_op_u16(Op::LOCAL_GET, width_slot, line);
                self.chunk().emit_op(Op::F64_MUL, line);
                self.chunk().emit_op_u16(Op::LOCAL_SET, offset_slot, line);

                // Advance AFTER computing the offset: this transfer goes where
                // the position said, and the next one goes after it.
                self.chunk().emit_op_u16(Op::LOCAL_GET, handle_slot, line);
                self.chunk().emit_string_const(F_POS, line);
                self.emit_handle_field(handle_slot, F_POS);
                self.chunk().emit_f64_const(1.0, line);
                self.chunk().emit_op(Op::F64_ADD, line);
                super::collections::emit_set(&mut self.chunks, current, line);
                self.chunk().emit_op(Op::DROP, line);
            }
            RecordAddress::Number(expr) => {
                // RELATIVE record numbers are 1-BASED in every language that
                // has them — COBOL `RELATIVE KEY`, VB `Get #f, n`, Fortran
                // `REC=n`. Record 1 is at offset 0.
                self.compile_expr(expr)?;
                self.chunk().emit_f64_const(1.0, line);
                self.chunk().emit_op(Op::F64_SUB, line);
                self.chunk().emit_op_u16(Op::LOCAL_GET, width_slot, line);
                self.chunk().emit_op(Op::F64_MUL, line);
                self.chunk().emit_op_u16(Op::LOCAL_SET, offset_slot, line);
            }
            RecordAddress::Key { .. } => unreachable!("rejected by compile_record_transfer"),
        }
        Ok(())
    }

    fn emit_record_write(
        &mut self,
        file: &Expression,
        record_type: &TypeRef,
        at: &RecordAddress,
        record: Option<&Expression>,
    ) -> Result<(), String> {
        let Some(record) = record else {
            return Err("RecordTransfer WRITE: no record to write".to_string());
        };
        let line = self.line;
        let current = self.current;

        let handle_slot = self.chunk().alloc_scratch(1);
        self.compile_expr(file)?;
        self.chunk().emit_op_u16(Op::LOCAL_SET, handle_slot, line);

        let desc_slot = self.chunk().alloc_scratch(1);
        self.emit_handle_field(handle_slot, F_DESC);
        self.chunk().emit_op_u16(Op::LOCAL_SET, desc_slot, line);

        let width_slot = self.chunk().alloc_scratch(1);
        self.emit_handle_field(handle_slot, F_WIDTH);
        self.chunk().emit_op_u16(Op::LOCAL_SET, width_slot, line);

        let offset_slot = self.chunk().alloc_scratch(1);
        self.emit_record_offset(handle_slot, width_slot, offset_slot, at)?;

        let text_slot = self.chunk().alloc_scratch(1);
        self.emit_record_bytes(record_type, record)?;
        self.chunk().emit_op_u16(Op::LOCAL_SET, text_slot, line);

        super::io::emit_write_descriptor_slot(
            &mut self.chunks[current],
            desc_slot,
            offset_slot,
            text_slot,
            line,
        );
        Ok(())
    }

    /// The record as the bytes it occupies: every member formatted to its own
    /// declared extent and concatenated, in declaration order.
    ///
    /// This is the whole of what a fixed-width record IS, and doing it here
    /// rather than in each walker is why `FieldStorage` exists.
    fn emit_record_bytes(
        &mut self,
        record_type: &TypeRef,
        record: &Expression,
    ) -> Result<(), String> {
        let layout = self
            .record_layout(record_type)
            .ok_or_else(|| {
                "RecordTransfer WRITE: the file's record type has no declared layout, so there \
                 is no way to say how wide each field is"
                    .to_string()
            })?;

        let line = self.line;
        let record_slot = self.chunk().alloc_scratch(1);
        self.compile_expr(record)?;
        self.chunk().emit_op_u16(Op::LOCAL_SET, record_slot, line);

        for (name, storage) in &layout {
            // CANONICAL key, not the declared spelling. A field name reaching
            // a property is subject to the language's own case rule — COBOL is
            // case-insensitive, so `EMP-NAME` is stored as `emp-name` — and
            // looking it up as declared reads a property that isn't there.
            // The symptom is not an error: every field comes back `undefined`
            // and is faithfully padded to its declared width, so the file is
            // exactly the right SIZE and entirely wrong.
            let key = self.canon(name);
            self.chunk().emit_op_u16(Op::LOCAL_GET, record_slot, line);
            self.chunk().emit_string_const(&key, line);
            let current = self.current;
            super::collections::emit_get(&mut self.chunks, current, line);
            self.emit_to_storage(*storage);
        }
        let current = self.current;
        super::strings::emit_concat(&mut self.chunks[current], layout.len(), line);
        Ok(())
    }

    /// One value → exactly `storage.bytes` characters. Stack: `[value]` → `[string]`.
    ///
    /// Numeric fields are right-justified and zero-filled, text fields are
    /// left-justified and space-filled — the two conventions every fixed-width
    /// record format shares, and the reason `Justify` is on the declaration
    /// rather than decided here.
    fn emit_to_storage(&mut self, storage: FieldStorage) {
        let line = self.line;
        let current = self.current;
        let width = storage.bytes as i64;

        // An implied decimal point is not stored, so the value is scaled to
        // integer digits before it is padded: `PIC 9(4)V99` holding 12.5 is
        // the six characters `001250`, not `12.50` in six columns.
        if storage.decimal_places > 0 {
            let scale = 10f64.powi(storage.decimal_places as i32);
            self.chunk().emit_f64_const(scale, line);
            self.chunk().emit_op(Op::F64_MUL, line);
            // Half-away-from-zero: COBOL `ROUNDED` and every fixed-width
            // record format round a stored midpoint away from zero, not to
            // even. The policy is stated because `emit_round` has three.
            super::math::emit_round(
                &mut self.chunks[current],
                vybe_ast::MidpointPolicy::HalfAwayFromZero,
                line,
            );
        }
        super::strings::emit_to_string(&mut self.chunks[current], line);
        self.chunk().emit_i32_const(width as i32, line);
        match storage.justify {
            Justify::Right => {
                self.chunk().emit_string_const("0", line);
                super::strings::emit_pad(
                    &mut self.chunks,
                    current,
                    3,
                    super::strings::PadSide::Start,
                    super::strings::CenterBias::Right,
                    line,
                );
            }
            Justify::Left => {
                self.chunk().emit_string_const(" ", line);
                super::strings::emit_pad(
                    &mut self.chunks,
                    current,
                    3,
                    super::strings::PadSide::End,
                    super::strings::CenterBias::Right,
                    line,
                );
            }
        }
        // A value WIDER than its field is truncated, not allowed to shift
        // every following field. Padding alone cannot shorten.
        self.chunk().emit_i32_const(0, line);
        self.chunk().emit_i32_const(width as i32, line);
        super::strings::emit_substring(&mut self.chunks[current], line);
    }

    /// Push `handle[field]`. Stack: `[]` → `[value]`.
    fn emit_handle_field(&mut self, handle_slot: u16, field: &str) {
        let line = self.line;
        let current = self.current;
        self.chunks[current].emit_op_u16(Op::LOCAL_GET, handle_slot, line);
        self.chunks[current].emit_string_const(field, line);
        super::collections::emit_get(&mut self.chunks, current, line);
    }

}

enum HandleInit {
    Slot(u16),
    Int(i64),
}

/// The type a `TypeRef` names, as written.
fn record_type_name(record: &TypeRef) -> Option<String> {
    match &record.kind {
        vybe_ast::TypeRefKind::Named { path, .. } => Some(path.display_name()),
        // A record type is a NAMED aggregate in every language that has record
        // files. An array, a tuple or a function type reaching here is a
        // walker error, not a shape to accommodate.
        _ => None,
    }
}
