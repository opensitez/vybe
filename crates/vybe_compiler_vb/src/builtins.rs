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
        for arg in args { self.compile_expression(arg)?; }
        self.emit_u8(Op::call, args.len() as u8);
        Ok(())
    }

    /// Try to compile a builtin function call. Returns Ok(Some(())) if handled.
    fn try_compile_builtin_call(&mut self, fname: &str, args: &[Expression]) -> Result<Option<()>, String> {
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
}
