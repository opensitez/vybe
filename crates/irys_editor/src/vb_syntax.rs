use egui::{text::LayoutJob, Color32, TextFormat, FontId};

/// irys Keywords for syntax highlighting
const irys_KEYWORDS: &[&str] = &[
    // Control flow
    "if", "then", "else", "elseif", "endif", "end", "select", "case",
    "for", "to", "step", "next", "each", "in", "while", "wend",
    "do", "loop", "until", "exit", "continue",
    // Declarations
    "dim", "redim", "preserve", "const", "static", "global",
    "public", "private", "friend", "option", "explicit", "base",
    "sub", "function", "property", "type", "enum", "declare",
    // Types and Object keywords
    "as", "new", "set", "let", "get", "typeof",
    "integer", "long", "single", "double", "string",
    "boolean", "byte", "currency", "date", "object",
    "variant", "any", "decimal",
    // Operators
    "and", "or", "not", "xor", "eqv", "imp",
    "mod", "is", "like", "to",
    // Parameters
    "byval", "byref", "optional", "paramarray",
    // Classes and modules
    "with", "withevents", "implements", "class", "module",
    "attribute", "event", "raiseevent",
    // Error handling
    "goto", "gosub", "return", "call", "raise",
    "resume", "on", "error", "line", "erl", "err",
    // Other
    "lib", "alias", "addressof", "open", "close", "input", "output",
    "binary", "random", "append", "read", "write", "seek",
    "def", "defint", "deflng", "defsng", "defdbl", "defstr",
    "defbool", "defvar", "defdate", "defbyte", "defdec", "defcur",
];

const irys_LITERALS: &[&str] = &[
    "true", "false", "nothing", "null", "empty",
];

const irys_BUILTINS: &[&str] = &[
    // Dialog functions
    "msgbox", "inputbox",
    // Output
    "print", "debug",
    // String functions
    "len", "left", "right", "mid", "trim", "ltrim", "rtrim",
    "ucase", "lcase", "instr", "instrrev", "replace", "split", "join",
    "space", "string", "strcomp", "strconv", "format",
    // Math functions
    "abs", "atn", "cos", "exp", "fix", "int", "log", "rnd", "sgn", "sin",
    "sqr", "tan", "round",
    // Date/Time functions
    "now", "date", "time", "timer", "dateadd", "datediff", "datepart",
    "dateserial", "timeserial", "year", "month", "day", "hour", "minute", "second",
    "weekday", "datevalue", "timevalue",
    // Conversion functions
    "cint", "clng", "csng", "cdbl", "cstr", "cbool", "cbyte", "cdate",
    "ccur", "cvar", "cdec", "val", "str", "hex", "oct",
    // Array functions
    "array", "lbound", "ubound", "redim", "erase",
    // File functions
    "eof", "lof", "loc", "freefile", "filedatetime", "fileattr",
    "filelen", "getattr", "setattr", "dir", "curdir", "chdir", "mkdir", "rmdir",
    "kill", "name", "filecopy",
    // Type checking
    "isnumeric", "isdate", "isempty", "isnull", "isobject", "isarray", "iserror",
    // Other functions
    "typename", "vartype", "rgb", "qbcolor", "choose", "switch", "iif",
    "createobject", "getobject", "loadpicture", "savepicture",
    "shell", "environ", "command", "inputb", "input",
    // Object references
    "me", "app", "screen", "printer", "clipboard", "forms", "err",
    "load", "unload", "doevents",
];

/// Create a syntax-highlighted LayoutJob for irys code
pub fn highlight_irys(code: &str, font_id: FontId) -> LayoutJob {
    let mut job = LayoutJob::default();

    for line in code.split('\n') {
        highlight_line(line, &mut job, font_id.clone());
        job.append("\n", 0.0, TextFormat {
            font_id: font_id.clone(),
            color: Color32::from_gray(200),
            ..Default::default()
        });
    }

    job
}

fn highlight_line(line: &str, job: &mut LayoutJob, font_id: FontId) {
    // Check if line is a comment
    let trimmed = line.trim_start();
    if trimmed.starts_with('\'') || trimmed.starts_with("Rem ") {
        // Entire line is a comment
        job.append(line, 0.0, TextFormat {
            font_id,
            color: Color32::from_rgb(87, 166, 74), // Green
            ..Default::default()
        });
        return;
    }

    // Split line into tokens
    let mut pos = 0;
    let line_lower = line.to_lowercase();

    while pos < line.len() {
        // Skip whitespace
        let ws_end = line[pos..].chars()
            .take_while(|c| c.is_whitespace())
            .map(|c| c.len_utf8())
            .sum::<usize>();

        if ws_end > 0 {
            job.append(&line[pos..pos + ws_end], 0.0, TextFormat {
                font_id: font_id.clone(),
                color: Color32::from_gray(200),
                ..Default::default()
            });
            pos += ws_end;
            continue;
        }

        // Check for comment
        if line[pos..].starts_with('\'') {
            job.append(&line[pos..], 0.0, TextFormat {
                font_id,
                color: Color32::from_rgb(87, 166, 74), // Green
                ..Default::default()
            });
            break;
        }

        // Check for string literal
        if line[pos..].starts_with('"') {
            let string_end = line[pos + 1..].find('"')
                .map(|i| i + 1)
                .unwrap_or(line.len() - pos - 1) + 1;
            let string_end = pos + string_end;

            job.append(&line[pos..string_end], 0.0, TextFormat {
                font_id: font_id.clone(),
                color: Color32::from_rgb(206, 145, 120), // Orange
                ..Default::default()
            });
            pos = string_end;
            continue;
        }

        // Extract word
        let word_end = line[pos..].find(|c: char| !c.is_alphanumeric() && c != '_')
            .unwrap_or(line.len() - pos);
        let word = &line[pos..pos + word_end];
        let word_lower = word.to_lowercase();

        // Determine color
        let color = if irys_KEYWORDS.contains(&word_lower.as_str()) {
            Color32::from_rgb(86, 156, 214) // Blue - keywords
        } else if irys_BUILTINS.contains(&word_lower.as_str()) {
            Color32::from_rgb(220, 220, 170) // Yellow - builtins
        } else if irys_LITERALS.contains(&word_lower.as_str()) {
            Color32::from_rgb(181, 206, 168) // Light green - literals
        } else if word.chars().next().map(|c| c.is_numeric()).unwrap_or(false) {
            Color32::from_rgb(181, 206, 168) // Light green - numbers
        } else {
            Color32::from_gray(220) // Default white/gray
        };

        job.append(word, 0.0, TextFormat {
            font_id: font_id.clone(),
            color,
            ..Default::default()
        });

        pos += word_end;

        // Handle remaining character (operator, punctuation)
        if pos < line.len() {
            let ch = line[pos..].chars().next().unwrap();
            if !ch.is_whitespace() {
                job.append(&line[pos..pos + ch.len_utf8()], 0.0, TextFormat {
                    font_id: font_id.clone(),
                    color: Color32::from_gray(200),
                    ..Default::default()
                });
                pos += ch.len_utf8();
            }
        }
    }
}

/// Extract all Sub/Function definitions from VB code
pub fn extract_procedures(code: &str) -> Vec<ProcedureInfo> {
    let mut procedures = Vec::new();

    for (line_num, line) in code.lines().enumerate() {
        let line_lower = line.trim().to_lowercase();

        // Check for Sub or Function declaration
        if line_lower.starts_with("sub ") ||
           line_lower.starts_with("private sub ") ||
           line_lower.starts_with("public sub ") ||
           line_lower.starts_with("function ") ||
           line_lower.starts_with("private function ") ||
           line_lower.starts_with("public function ") {

            // Extract the procedure name
            let parts: Vec<&str> = line.split_whitespace().collect();

            // Find the name (after Sub/Function keyword)
            let mut name_idx = 0;
            for (i, part) in parts.iter().enumerate() {
                let part_lower = part.to_lowercase();
                if part_lower == "sub" || part_lower == "function" {
                    name_idx = i + 1;
                    break;
                }
            }

            if name_idx < parts.len() {
                let name_with_params = parts[name_idx];
                // Remove parameters if present
                let name = if let Some(paren_pos) = name_with_params.find('(') {
                    &name_with_params[..paren_pos]
                } else {
                    name_with_params
                };

                let proc_type = if line_lower.contains("function") {
                    ProcedureType::Function
                } else {
                    ProcedureType::Sub
                };

                procedures.push(ProcedureInfo {
                    name: name.to_string(),
                    line: line_num,
                    proc_type,
                });
            }
        }
    }

    procedures
}

#[derive(Debug, Clone)]
pub struct ProcedureInfo {
    pub name: String,
    pub line: usize,
    pub proc_type: ProcedureType,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ProcedureType {
    Sub,
    Function,
}

impl ProcedureType {
    pub fn icon(&self) -> &str {
        match self {
            ProcedureType::Sub => "📋",
            ProcedureType::Function => "⚙️",
        }
    }
}
