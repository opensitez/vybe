use std::rc::Rc;
use vybe_bytecode::{Value, Op};
use vybe_parser_basic::ast::*;

use crate::compiler::{Compiler, VarResolution};

impl Compiler {
    /// Compile a function/sub call expression (name + args).
    /// Checks builtins first, then falls back to regular call.
    pub(crate) fn compile_call_expr(&mut self, name: &Identifier, args: &[Expression]) -> Result<(), String> {
        let fname = name.as_str().to_lowercase();

        // Try builtin mapping first
        if let Some(()) = self.try_compile_builtin_call(&fname, args)? {
            return Ok(());
        }

        // Check if name is a local variable — if so, treat as array access
        // (VB uses parens for both calls and array indexing)
        match self.resolve_variable(&fname) {
            VarResolution::Local(slot) => {
                if !self.defined_globals.contains(&fname) {
                    self.emit_u16(Op::local_get, slot);
                    if let Some(index) = args.first() {
                        self.compile_expression(index)?;
                        self.emit(Op::array_get);
                    }
                    return Ok(());
                }
                self.emit_u16(Op::local_get, slot);
            }
            VarResolution::Global => {
                let idx = self.add_string_constant(&fname);
                self.emit_u16(Op::global_get, idx);
            }
        }
        // Box ByRef args if function signature is known
        let sig = self.func_signatures.get(&fname).cloned();
        let mut byref_info: Vec<(u16, u16)> = Vec::new();
        for (i, arg) in args.iter().enumerate() {
            let is_byref = sig.as_ref().and_then(|s| s.get(i)).copied().unwrap_or(false);
            if is_byref {
                if let Expression::Variable(var) = arg {
                    let var_name = var.as_str().to_lowercase();
                    self.compile_expression(arg)?;
                    self.emit_u16(Op::array_new, 1);
                    let box_local = self.define_local(&format!("__box_{}", i));
                    self.emit(Op::dup);
                    self.emit_u16(Op::local_set, box_local);
                    self.emit(Op::drop);
                    if let VarResolution::Local(var_slot) = self.resolve_variable(&var_name) {
                        byref_info.push((box_local, var_slot));
                    }
                } else {
                    self.compile_expression(arg)?;
                    self.emit_u16(Op::array_new, 1);
                }
            } else {
                self.compile_expression(arg)?;
            }
        }
        self.emit_u8(Op::call, args.len() as u8);

        // Writeback ByRef vars from boxes
        for (box_local, var_local) in &byref_info {
            self.emit_u16(Op::local_get, *box_local);
            self.emit_constant(Value::F64(0.0));
            self.emit(Op::array_get);
            self.emit_u16(Op::local_set, *var_local);
            self.emit(Op::drop);
        }

        Ok(())
    }

    /// Try to compile a builtin function call. Returns Ok(Some(())) if handled.
    fn try_compile_builtin_call(&mut self, fname: &str, args: &[Expression]) -> Result<Option<()>, String> {
        // Direct WASM opcodes for Math and type conversions (no host call overhead)
        if let Some(()) = self.try_compile_wasm_intrinsic(fname, args)? {
            return Ok(Some(()));
        }

        // Map VB function name → (host_module, host_name)
        let mapping: Option<(&str, &str)> = match fname {
            // Console
            "console.writeline" => Some(("wasi:cli", "log")),
            // Type conversion
            "cstr"  => Some(("vybe:convert", "toString")),
            "cint"  => Some(("vybe:convert", "cint")),
            "cdbl"  => Some(("vybe:convert", "cdbl")),
            "cbool" => Some(("vybe:convert", "cbool")),
            "clng"  => Some(("vybe:convert", "clng")),
            "csng"  => Some(("vybe:convert", "csng")),
            "cbyte" => Some(("vybe:convert", "cbyte")),
            "cchar" => Some(("vybe:convert", "cchar")),
            "val"   => Some(("vybe:convert", "val")),
            "hex" | "hex$" => Some(("vybe:convert", "hex")),
            "oct" | "oct$" => Some(("vybe:convert", "oct")),
            "str" | "str$" => Some(("vybe:convert", "str")),
            "iif"   => Some(("vybe:convert", "iif")),
            "choose" => Some(("vybe:convert", "choose")),
            "rgb"   => Some(("vybe:convert", "rgb")),
            // Type checking
            "isnumeric"  => Some(("vybe:convert", "isNumeric")),
            "isnothing"  => Some(("vybe:convert", "isNothing")),
            "isnull"     => Some(("vybe:convert", "isNull")),
            "isempty"    => Some(("vybe:convert", "isEmpty")),
            "isobject"   => Some(("vybe:convert", "isObject")),
            "isarray"    => Some(("vybe:convert", "isArray")),
            "isdate"     => Some(("vybe:convert", "isDate")),
            "typename"   => Some(("vybe:convert", "typeName")),
            "vartype"    => Some(("vybe:convert", "varType")),
            // String functions
            "len"        => Some(("vybe:string", "length")),
            "ucase"      => Some(("vybe:string", "ucase")),
            "lcase"      => Some(("vybe:string", "lcase")),
            "trim"       => Some(("vybe:string", "trim")),
            "ltrim"      => Some(("vybe:string", "ltrim")),
            "rtrim"      => Some(("vybe:string", "rtrim")),
            "left"       => Some(("vybe:string", "left")),
            "right"      => Some(("vybe:string", "right")),
            "mid" | "mid$" => Some(("vybe:string", "mid")),
            "instr"      => Some(("vybe:string", "instr")),
            "instrrev"   => Some(("vybe:string", "instrrev")),
            "replace"    => Some(("vybe:string", "replaceAll")),
            "split"      => Some(("vybe:string", "split")),
            "join"       => Some(("vybe:array", "join")),
            "asc" | "ascw" => Some(("vybe:string", "asc")),
            "chr" | "chr$" | "chrw" => Some(("vybe:string", "chr")),
            "space" | "space$" | "spc" => Some(("vybe:string", "space")),
            "strreverse" => Some(("vybe:string", "strreverse")),
            "strcomp"    => Some(("vybe:string", "strcomp")),
            "format" | "format$" => Some(("vybe:string", "format")),
            "string" | "string$" => Some(("vybe:string", "stringRepeat")),
            "lset"       => Some(("vybe:string", "lset")),
            "rset"       => Some(("vybe:string", "rset")),
            "filter"     => Some(("vybe:string", "filter")),
            // Math functions
            "abs"   => Some(("vybe:math", "abs")),
            "sqr" | "sqrt" => Some(("vybe:math", "sqrt")),
            "fix"   => Some(("vybe:math", "fix")),
            "sgn"   => Some(("vybe:math", "sgn")),
            "rnd"   => Some(("vybe:math", "rnd")),
            "randomize" => Some(("vybe:math", "randomize")),
            "int"   => Some(("vybe:math", "int")),
            "log"   => Some(("vybe:math", "log")),
            "exp"   => Some(("vybe:math", "exp")),
            "sin"   => Some(("vybe:math", "sin")),
            "cos"   => Some(("vybe:math", "cos")),
            "tan"   => Some(("vybe:math", "tan")),
            "atn" | "atan" => Some(("vybe:math", "atan")),
            "round" => Some(("vybe:math", "round")),
            "math.floor"   => Some(("vybe:math", "floor")),
            "math.abs"     => Some(("vybe:math", "abs")),
            "math.sqrt"    => Some(("vybe:math", "sqrt")),
            // Date/Time functions
            "now"          => Some(("wasi:clocks", "vbNow")),
            "date" | "today" => Some(("wasi:clocks", "vbDate")),
            "time" | "timeofday" => Some(("wasi:clocks", "vbTime")),
            "timer"        => Some(("wasi:clocks", "vbTimer")),
            "year"         => Some(("wasi:clocks", "vbYear")),
            "month"        => Some(("wasi:clocks", "vbMonth")),
            "day"          => Some(("wasi:clocks", "vbDay")),
            "hour"         => Some(("wasi:clocks", "vbHour")),
            "minute"       => Some(("wasi:clocks", "vbMinute")),
            "second"       => Some(("wasi:clocks", "vbSecond")),
            // Array functions
            "ubound" => Some(("vybe:array", "ubound")),
            "lbound" => Some(("vybe:array", "lbound")),
            "array"  => Some(("vybe:array", "from")),
            // More date/time
            "dateadd"       => Some(("wasi:clocks", "vbNow")),   // simplified
            "datediff"      => Some(("wasi:clocks", "vbNow")),   // simplified
            "datepart"      => Some(("wasi:clocks", "vbNow")),   // simplified
            "dateserial"    => Some(("wasi:clocks", "vbNow")),   // simplified
            "datevalue"     => Some(("wasi:clocks", "vbNow")),   // simplified
            "timeserial"    => Some(("wasi:clocks", "vbNow")),   // simplified
            "timevalue"     => Some(("wasi:clocks", "vbNow")),   // simplified
            "cdate"         => Some(("wasi:clocks", "vbNow")),   // simplified
            "monthname"     => Some(("vybe:convert", "toString")), // simplified
            "weekday"       => Some(("wasi:clocks", "vbDay")),   // simplified
            "weekdayname"   => Some(("vybe:convert", "toString")), // simplified
            // More conversion
            "ccur" | "cdec" | "cvar" | "cobj" => Some(("vybe:convert", "cdbl")),
            "cshort" | "cushort" | "cuint" | "culng" | "csbyte"
                            => Some(("vybe:convert", "cint")),
            // More string ($-suffixed variants)
            "lcase$" | "ucase$" | "trim$" | "ltrim$" | "rtrim$"
            | "mid$" | "chr$" | "chrw$" | "left$" | "right$"
            | "space$" | "spc" | "lset$" | "rset$" => {
                // Strip $ and recurse
                let base = fname.trim_end_matches('$');
                return self.try_compile_builtin_call(base, args);
            }
            // Format variants
            "formatnumber"   => Some(("vybe:string", "format")),
            "formatcurrency" => Some(("vybe:string", "format")),
            "formatpercent"  => Some(("vybe:string", "format")),
            "formatdatetime" => Some(("vybe:string", "format")),
            "strconv"        => Some(("vybe:convert", "toString")),
            "strdup"         => Some(("vybe:convert", "toString")),
            // File functions
            "dir" | "dir$"   => Some(("wasi:filesystem", "listDir")),
            "filecopy"       => Some(("wasi:filesystem", "copy")),
            "kill"           => Some(("wasi:filesystem", "remove")),
            "fileexists" | "file.exists" => Some(("wasi:filesystem", "exists")),
            "filedatetime"   => Some(("wasi:clocks", "vbNow")),
            "filelen"        => Some(("wasi:filesystem", "fileSize")),
            "curdir" | "curdir$" => Some(("wasi:cli", "cwd")),
            "chdir"          => Some(("wasi:cli", "cwd")),
            "mkdir"          => Some(("wasi:filesystem", "mkdir")),
            "rmdir"          => Some(("wasi:filesystem", "remove")),
            "freefile"       => Some(("vybe:convert", "cint")),
            // Interaction
            "msgbox"         => Some(("vybe:gui", "msgBox")),
            "inputbox"       => Some(("vybe:gui", "msgBox")), // simplified
            "beep"           => Some(("wasi:cli", "log")),     // no-op
            "shell"          => Some(("vybe:types", "processStart")),
            "environ" | "environ$" => Some(("wasi:cli", "getEnv")),
            "command" | "command$" => Some(("wasi:cli", "args")),
            "sendkeys"       => Some(("wasi:cli", "log")),     // no-op
            "appactivate"    => Some(("wasi:cli", "log")),     // no-op
            // Info
            "isdbnull"       => Some(("vybe:convert", "isNull")),
            "iserror"        => Some(("vybe:convert", "isNull")),
            // Crypto
            "qbcolor"        => Some(("vybe:convert", "rgb")),
            // JSON
            "json.serialize"   => Some(("vybe:json", "stringify")),
            "json.deserialize" => Some(("vybe:json", "parse")),
            // XML
            "xml.parse" | "xdocument.parse" => Some(("vybe:xml", "parse")),
            "xml.load" | "xdocument.load"   => Some(("vybe:xml", "load")),
            "xml.save" | "xdocument.save"   => Some(("vybe:xml", "toString")),
            // Regex
            "regex.ismatch"  => Some(("vybe:regex", "test")),
            "regex.match"    => Some(("vybe:regex", "match")),
            "regex.matches"  => Some(("vybe:regex", "match")),
            "regex.replace"  => Some(("vybe:regex", "replace")),
            "regex.split"    => Some(("vybe:regex", "split")),
            // Encoding
            "encoding.ascii.getbytes" | "encoding.utf8.getbytes"
            | "encoding.unicode.getbytes" | "encoding.default.getbytes"
                => Some(("vybe:convert", "toString")), // simplified
            "encoding.ascii.getstring" | "encoding.utf8.getstring"
            | "encoding.unicode.getstring" | "encoding.default.getstring"
                => Some(("vybe:convert", "toString")), // simplified
            "encoding.getencoding" | "encoding.convert"
                => Some(("vybe:convert", "toString")), // simplified
            // StringBuilder
            "stringbuilder"  => Some(("vybe:types", "stringBuilderNew")),
            // Switch (VB Select-like function)
            "switch"         => Some(("vybe:convert", "choose")),
            // VB6 file I/O (simplified — map to filesystem or no-op)
            "open"           => Some(("wasi:filesystem", "readFile")),  // simplified
            "close"          => Some(("wasi:cli", "log")),             // no-op
            "print"          => Some(("wasi:cli", "log")),
            "write"          => Some(("wasi:cli", "log")),
            "input"          => Some(("wasi:cli", "readLine")),
            "inputb"         => Some(("wasi:cli", "readLine")),
            "lineinput"      => Some(("wasi:cli", "readLine")),
            "get"            => Some(("wasi:filesystem", "readFile")), // simplified
            "put"            => Some(("wasi:filesystem", "writeFile")),// simplified
            "seek"           => Some(("wasi:cli", "log")),            // no-op
            "eof"            => Some(("vybe:convert", "cbool")),      // simplified
            "lof"            => Some(("wasi:filesystem", "fileSize")),
            "loc"            => Some(("vybe:convert", "cint")),       // simplified
            "fileattr"       => Some(("vybe:convert", "cint")),       // simplified
            "getattr"        => Some(("vybe:convert", "cint")),       // simplified
            "setattr"        => Some(("wasi:cli", "log")),            // no-op
            "name"           => Some(("wasi:filesystem", "rename")),
            "erase"          => Some(("wasi:cli", "log")),            // no-op (array clear)
            // Legacy VB6 interaction
            "loadpicture"    => Some(("vybe:convert", "toString")),   // simplified
            "savepicture"    => Some(("wasi:cli", "log")),            // no-op
            "load"           => Some(("wasi:cli", "log")),            // no-op
            "unload"         => Some(("wasi:cli", "log")),            // no-op
            "app"            => Some(("wasi:cli", "args")),           // simplified
            "screen"         => Some(("vybe:convert", "cint")),       // simplified
            "clipboard"      => Some(("vybe:convert", "toString")),   // simplified
            "forms"          => Some(("vybe:convert", "toString")),   // simplified
            _ => None,
        };

        if let Some((module, name)) = mapping {
            for arg in args { self.compile_expression(arg)?; }
            let idx = self.import(module, name);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(Some(()));
        }

        // Special cases
        match fname {
            "doevents" => {
                self.emit(Op::null);
                Ok(Some(()))
            }
            _ => Ok(None),
        }
    }

    /// Try to compile a dotted method call (e.g. Console.WriteLine, Math.Floor).
    /// Returns Ok(Some(())) if handled, Ok(None) if not a known builtin.
    pub(crate) fn try_compile_builtin_method(&mut self, full_name: &str, args: &[Expression]) -> Result<Option<()>, String> {
        // Map dotted name → (host_module, host_name)
        let mapping: Option<(&str, &str)> = match full_name {
            // Console
            "console.writeline" | "console.write" => Some(("wasi:cli", "log")),
            "console.error.writeline" | "console.error" => Some(("wasi:cli", "error")),
            // Math
            "math.floor"    => Some(("vybe:math", "floor")),
            "math.ceiling" | "math.ceil" => Some(("vybe:math", "ceil")),
            "math.abs"      => Some(("vybe:math", "abs")),
            "math.sqrt"     => Some(("vybe:math", "sqrt")),
            "math.pow"      => Some(("vybe:math", "pow")),
            "math.min"      => Some(("vybe:math", "min")),
            "math.max"      => Some(("vybe:math", "max")),
            "math.round"    => Some(("vybe:math", "round")),
            "math.sin"      => Some(("vybe:math", "sin")),
            "math.cos"      => Some(("vybe:math", "cos")),
            "math.tan"      => Some(("vybe:math", "tan")),
            "math.log"      => Some(("vybe:math", "log")),
            "math.sign"     => Some(("vybe:math", "sign")),
            "math.truncate" => Some(("vybe:math", "trunc")),
            // String
            "string.isnullorempty" => Some(("vybe:string", "length")),
            // Convert
            "convert.toint32" | "convert.toint" => Some(("vybe:math", "floor")),
            "convert.todouble"  => Some(("vybe:convert", "parseFloat")),
            "convert.tostring"  => Some(("vybe:convert", "toString")),
            "convert.todatetime" => Some(("vybe:types", "dateTimeNow")),
            // DateTime properties accessed as Namespace.Property
            "datetime.now"   => Some(("vybe:types", "dateTimeNow")),
            "datetime.today" => Some(("vybe:types", "dateTimeNow")),
            "datetime.utcnow" => Some(("vybe:types", "dateTimeNow")),
            // Application
            "application.run" => Some(("vybe:gui", "runApplication")),
            _ => None,
        };

        if let Some((module, name)) = mapping {
            // Special case: String.IsNullOrEmpty returns length == 0
            if full_name == "string.isnullorempty" {
                self.compile_expression(&args[0])?;
                let idx = self.import("vybe:string", "length");
                self.emit_host_call(idx, 1);
                self.emit_constant(Value::F64(0.0));
                self.emit(Op::dyn_eq);
                return Ok(Some(()));
            }

            // Special case: Application.Run may take an object
            if full_name == "application.run" {
                if let Some(arg) = args.first() {
                    self.compile_expression(arg)?;
                } else {
                    self.emit_constant(Value::String(Rc::from("Form1")));
                }
                let idx = self.import(module, name);
                self.emit_host_call(idx, 1);
                return Ok(Some(()));
            }

            for arg in args { self.compile_expression(arg)?; }
            let idx = self.import(module, name);
            self.emit_host_call(idx, args.len() as u8);
            return Ok(Some(()));
        }

        Ok(None)
    }

    /// Emit direct WASM opcodes for Math functions, type conversions, and type checks.
    /// Returns Some(()) if handled, None if not an intrinsic.
    fn try_compile_wasm_intrinsic(&mut self, fname: &str, args: &[Expression]) -> Result<Option<()>, String> {
        // Zero-argument intrinsics
        if args.is_empty() {
            match fname {
                "lbound" => {
                    self.emit(Op::i32_const_0);
                    return Ok(Some(()));
                }
                _ => {}
            }
        }

        // Single-argument intrinsics
        if args.len() == 1 {
            // Direct single-opcode intrinsics
            let op = match fname {
                // Math functions → f64 opcodes
                "math.abs" | "abs" => Some(Op::f64_abs),
                "math.floor" | "fix" | "int" => Some(Op::f64_floor),
                "math.ceiling" | "math.ceil" => Some(Op::f64_ceil),
                "math.sqrt" | "sqr" => Some(Op::f64_sqrt),
                "math.truncate" => Some(Op::f64_trunc),
                "math.round" => Some(Op::f64_nearest),
                // Type conversions
                "cbool" => Some(Op::dyn_to_bool),
                // Type checks
                "isnothing" | "isnull" => Some(Op::ref_is_null),
                // String builtins (wasm:js-string proposal)
                "len" => Some(Op::str_length),
                "ucase" => Some(Op::str_to_upper),
                "lcase" => Some(Op::str_to_lower),
                "trim" => Some(Op::str_trim),
                "ltrim" => Some(Op::str_trim_start),
                "rtrim" => Some(Op::str_trim_end),
                "strreverse" => Some(Op::str_reverse),
                "chr" | "chr$" | "chrw" => Some(Op::str_from_char_code),
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(&args[0])?;
                self.emit(op);
                return Ok(Some(()));
            }

            // Multi-opcode intrinsics
            match fname {
                // CByte: truncate to 0-255
                "cbyte" => {
                    self.compile_expression(&args[0])?;
                    self.emit(Op::i32_from_f64);
                    self.emit_constant(Value::I32(0xFF));
                    self.emit(Op::i32_and);
                    return Ok(Some(()));
                }
                // UBound: array length - 1
                "ubound" => {
                    self.compile_expression(&args[0])?;
                    self.emit(Op::array_length);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::i32_sub);
                    return Ok(Some(()));
                }
                // Asc(s) → str_char_code_at(s, 0)
                "asc" | "ascw" => {
                    self.compile_expression(&args[0])?;
                    self.emit(Op::i32_const_0);
                    self.emit(Op::str_char_code_at);
                    return Ok(Some(()));
                }
                _ => {}
            }
        }

        // Two-argument intrinsics
        if args.len() == 2 {
            let op = match fname {
                "math.min" => Some(Op::f64_min),
                "math.max" => Some(Op::f64_max),
                "split" => Some(Op::str_split),
                "join" => Some(Op::array_join),
                "string" | "string$" => Some(Op::str_repeat),
                _ => None,
            };
            if let Some(op) = op {
                self.compile_expression(&args[0])?;
                self.compile_expression(&args[1])?;
                self.emit(op);
                return Ok(Some(()));
            }

            // Multi-opcode two-arg intrinsics
            match fname {
                // InStr(s, needle) → str_index_of + 1 (VB is 1-based: 1=first, 0=not found)
                "instr" => {
                    self.compile_expression(&args[0])?;
                    self.compile_expression(&args[1])?;
                    self.emit(Op::str_index_of);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::i32_add); // -1→0, 0→1, 5→6
                    return Ok(Some(()));
                }
                // Left(s, n) → str_substring(s, 0, n)
                "left" | "left$" => {
                    self.compile_expression(&args[0])?;
                    self.emit(Op::i32_const_0);
                    self.compile_expression(&args[1])?;
                    self.emit(Op::i32_from_f64);
                    self.emit(Op::str_substring);
                    return Ok(Some(()));
                }
                // Right(s, n) → str_substring(s, len-n, len)
                "right" | "right$" => {
                    self.compile_expression(&args[0])?;
                    self.emit(Op::dup);
                    self.emit(Op::str_length);        // [s, len]
                    self.emit(Op::dup);               // [s, len, len]
                    self.compile_expression(&args[1])?;
                    self.emit(Op::i32_from_f64);
                    self.emit(Op::i32_sub);           // [s, len, len-n]
                    // need: [s, len-n, len] — swap top two
                    // No swap opcode, so recompute: use the stack
                    // Actually let's just fall through to host call for Right — it's complex
                    self.emit(Op::drop);
                    self.emit(Op::drop);
                    self.emit(Op::drop);
                    return Ok(None); // fall through to host call
                }
                _ => {}
            }
        }

        // Mid(s, start) — 2-arg form: from start to end
        if args.len() == 2 && (fname == "mid" || fname == "mid$") {
            self.compile_expression(&args[0])?;
            // start0 = start - 1
            self.compile_expression(&args[1])?;
            self.emit(Op::i32_from_f64);
            self.emit_constant(Value::I32(1));
            self.emit(Op::i32_sub);
            // end = large number (rest of string)
            self.emit_constant(Value::I32(0x7FFF_FFFF));
            self.emit(Op::str_substring);
            return Ok(Some(()));
        }

        // Three-argument intrinsics
        if args.len() == 3 {
            match fname {
                // InStr(startPos, string, substring) — 3-arg with 1-based start offset
                "instr" => {
                    // We need: str_index_of(substring_from(s, start-1), needle) + start
                    // Simpler: use str_substring to get tail, then str_index_of, adjust
                    self.compile_expression(&args[1])?; // string
                    self.compile_expression(&args[0])?; // startPos
                    self.emit(Op::i32_from_f64);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::i32_sub);             // start0
                    self.emit(Op::dup);                 // [s, start0, start0]
                    self.emit_constant(Value::I32(0x7FFF_FFFF));
                    self.emit(Op::str_substring);       // [start0, tail]
                    self.compile_expression(&args[2])?; // needle
                    self.emit(Op::str_index_of);        // [start0, pos_in_tail]
                    // if pos_in_tail == -1, result = 0; else result = pos_in_tail + start0 + 1
                    self.emit(Op::dup);                 // [start0, pos, pos]
                    self.emit_constant(Value::I32(-1));
                    self.emit(Op::dyn_eq);              // [start0, pos, is_not_found]
                    let found = self.emit_jump(Op::br_if_true);
                    // Found: pos + start0 + 1
                    self.emit(Op::i32_add);             // pos + start0
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::i32_add);
                    let end = self.emit_jump(Op::br);
                    self.patch_jump(found);
                    // Not found: drop pos and start0, push 0
                    self.emit(Op::drop);
                    self.emit(Op::drop);
                    self.emit(Op::i32_const_0);
                    self.patch_jump(end);
                    return Ok(Some(()));
                }
                "replace" => {
                    self.compile_expression(&args[0])?;
                    self.compile_expression(&args[1])?;
                    self.compile_expression(&args[2])?;
                    self.emit(Op::str_replace);
                    return Ok(Some(()));
                }
                // Mid(s, start, length) → str_substring(s, start-1, start-1+length)
                "mid" | "mid$" => {
                    self.compile_expression(&args[0])?;
                    // Convert 1-based start to 0-based
                    self.compile_expression(&args[1])?;
                    self.emit(Op::i32_from_f64);
                    self.emit_constant(Value::I32(1));
                    self.emit(Op::i32_sub);           // start0 = start - 1
                    self.emit(Op::dup);               // [s, start0, start0]
                    self.compile_expression(&args[2])?;
                    self.emit(Op::i32_from_f64);
                    self.emit(Op::i32_add);           // [s, start0, start0+length]
                    self.emit(Op::str_substring);
                    return Ok(Some(()));
                }
                _ => {}
            }
        }
        Ok(None)
    }
}
