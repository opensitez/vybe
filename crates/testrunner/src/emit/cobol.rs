//! COBOL emitter: one extracted case → a standalone `.cob` test.
//!
//! Two shapes, and the first is the majority:
//!
//! **`compile_ok` (2,423 of 3,844)** carried no expected output. Rather than
//! emitting them compile-only they are emitted as ordinary RUNS with no
//! assertion — "runs to completion, exit 0" subsumes "the frontend accepted
//! it" and additionally catches runtime failures that compile mode cannot see.
//!
//! **`run_prints` (1,471)** pairs the i-th `DISPLAY` with the i-th expected
//! line. COBOL has no parameterised paragraph, so there is nothing to
//! `PERFORM` — the check is emitted inline after each `DISPLAY`. See
//! `harness/cobol/check.cob` for the shape. Failure is signalled BOTH ways —
//! `MOVE 1 TO RETURN-CODE` and `RAISE EXCEPTION EC-PROGRAM` — because each
//! runtime honours exactly one: cobc exits 1 on the RETURN-CODE and treats
//! RAISE as an unimplemented no-op, while Vybe throws on the RAISE and ignores
//! RETURN-CODE (no exit-code path exists in the VM). `STOP RUN WITH ERROR
//! STATUS` is NOT in the grammar — using it made the check a PARSE ERROR, so
//! every test "failed" for the wrong reason.
//!
//! The program is assembled from two sources — `p(data, body)` builds
//! WORKING-STORAGE and PROCEDURE DIVISION separately — so the data division
//! arrives in `case.prelude`.

use crate::emit::go::Pairing;
use crate::extract::Case;

pub struct Emitted {
    pub text: String,
    pub pairing: Pairing }

pub fn emit(case: &Case, origin: &str, slug: &str, _harness: &str) -> Emitted {
    // `*>` is the FREE-format comment. The fixed-format `      *` in column 7
    // is rejected by both `cobc -free` and Vybe, which is what the corpus uses.
    let header = format!("*> vybe-test: {slug}\n*> origin: {origin}\n");

    let Some(expected) = case.expected.as_ref() else {
        // Run it, do not merely compile it.
        return Emitted {
            text: format!("{header}{}\n", assemble(&case.source, case.prelude.as_deref(), false)),
            pairing: Pairing::Direct };
    };

    let displays = find_displays(&case.source);
    if let Some(reason) = unpairable(&case.source, &displays, expected) {
        // RUNTIME pairing: an expected TABLE plus a counter, checked where each
        // DISPLAY stands. The counter advances as the program runs, so a
        // DISPLAY inside a PERFORM pairs correctly and a conditional branch
        // contributes exactly the line it actually produced — 165 loop cases
        // and 116 count mismatches that static pairing can only refuse.
        if let Some((body, table)) = runtime_paired(&case.source, &displays, expected, &reason) {
            return Emitted {
                text: format!(
                    "{header}{}\n",
                    assemble_with(&body, case.prelude.as_deref(), true, &table)
                ),
                pairing: Pairing::Direct };
        }
        return Emitted {
            text: format!("{header}{}\n", assemble(&case.source, case.prelude.as_deref(), false)),
            pairing: Pairing::Unpairable(reason) };
    }

    let mut body = case.source.clone();
    for (i, (_, end, operands)) in displays.iter().enumerate().rev() {
        // Concatenate the operand list into one field and compare that. This
        // is what DISPLAY itself produces, so it works for a single operand, a
        // multi-operand list and a bare literal alike — per-operand pairing
        // could only ever express the first.
        //
        // `DELIMITED SIZE` is repeated after EVERY operand on purpose: Vybe
        // does not propagate a single trailing delimiter back over the
        // preceding operands the way cobc does, so the explicit form is the
        // one both runtimes agree on (measured).
        let sources = operands
            .iter()
            .map(|o| format!("{o} DELIMITED SIZE"))
            .collect::<Vec<_>>()
            .join(" ");
        let want = expected[i].replace('"', "'");
        let check = format!(
            "\n    MOVE SPACES TO WS-VYBE-L\n    STRING {sources} INTO WS-VYBE-L\n    \
             IF WS-VYBE-L NOT = \"{want}\"\n        \
             DISPLAY \"FAIL: want [{want}] got [\" WS-VYBE-L \"]\"\n        \
             MOVE 1 TO RETURN-CODE\n        RAISE EXCEPTION EC-PROGRAM\n    END-IF."
        );
        body.insert_str(*end, &check);
    }

    Emitted {
        text: format!(
            "{header}{}\n",
            assemble(&body, case.prelude.as_deref(), true)
        ),
        pairing: Pairing::Direct }
}

/// Build the runtime-paired body plus the WORKING-STORAGE table it needs.
///
/// Only the reasons that are about PAIRING can be rescued this way. A DISPLAY
/// inside an `ON SIZE ERROR`-style clause is not a complete statement, so no
/// check can be appended after it however the lines are matched up.
fn runtime_paired(
    src: &str,
    displays: &[Display],
    expected: &[String],
    reason: &str,
) -> Option<(String, String)> {
    if displays.is_empty() || expected.is_empty() || reason.contains("clause") {
        return None;
    }
    let n = expected.len();
    // An EVALUATE on the counter, NOT a table. `01 WS-VYBE-W` + FILLER VALUEs +
    // `REDEFINES … OCCURS` is the textbook COBOL expected-value table, and it
    // does not work under Vybe: the subscripted read came back unequal to a
    // literal it held verbatim (measured — `got [02]` against an expectation of
    // exactly "02"). EVALUATE needs no REDEFINES, no OCCURS and no subscript,
    // only the literal comparison the static path already proves out.
    let mut table = String::from("01 WS-VYBE-I PIC 9(4) VALUE 0.");
    if false {
        table.clear();
    }

    let mut body = src.to_string();
    for (_, end, operands) in displays.iter().rev() {
        let sources = operands
            .iter()
            .map(|o| format!("{o} DELIMITED SIZE"))
            .collect::<Vec<_>>()
            .join(" ");
        let mut arms = String::new();
        for (i, want) in expected.iter().enumerate() {
            let want = want.replace('"', "'");
            arms.push_str(&format!(
                "\n        WHEN {}\n            IF WS-VYBE-L NOT = \"{want}\"\n                \
                 DISPLAY \"FAIL at {} want [{want}] got [\" WS-VYBE-L \"]\"\n                \
                 MOVE 1 TO RETURN-CODE\n                RAISE EXCEPTION EC-PROGRAM\n            END-IF",
                i + 1,
                i + 1
            ));
        }
        let check = format!(
            "\n    ADD 1 TO WS-VYBE-I\n    MOVE SPACES TO WS-VYBE-L\n    \
             STRING {sources} INTO WS-VYBE-L\n    EVALUATE WS-VYBE-I{arms}\n        \
             WHEN OTHER\n            DISPLAY \"FAIL: more than {n} line(s)\"\n            \
             MOVE 1 TO RETURN-CODE\n            RAISE EXCEPTION EC-PROGRAM\n    END-EVALUATE."
        );
        body.insert_str(*end, &check);
    }
    // Too FEW lines is a failure too, and only the end can see it.
    body.push_str(&format!(
        "\n    IF WS-VYBE-I NOT = {n}\n        \
         DISPLAY \"FAIL: \" WS-VYBE-I \" line(s), wanted {n}\"\n        \
         MOVE 1 TO RETURN-CODE\n        RAISE EXCEPTION EC-PROGRAM\n    END-IF."
    ));
    Some((body, table))
}

/// Split a DISPLAY operand list on whitespace OUTSIDE quotes — a literal may
/// contain spaces (`DISPLAY "lit " WS-A`).
fn split_operands(text: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut quote: Option<char> = None;
    for ch in text.chars() {
        match quote {
            Some(q) => {
                cur.push(ch);
                if ch == q {
                    quote = None;
                }
            }
            None => match ch {
                '"' | '\'' => {
                    quote = Some(ch);
                    cur.push(ch);
                }
                c if c.is_whitespace() => {
                    if !cur.is_empty() {
                        out.push(std::mem::take(&mut cur));
                    }
                }
                c => cur.push(c) } }
    }
    if !cur.is_empty() {
        out.push(cur);
    }
    out
}

/// Rebuild what the test's `p(data, body)` helper produced.
fn assemble(body: &str, prelude: Option<&str>, needs_scratch: bool) -> String {
    assemble_with(body, prelude, needs_scratch, "")
}

fn assemble_with(
    body: &str,
    prelude: Option<&str>,
    needs_scratch: bool,
    extra_data: &str,
) -> String {
    // Some cases carry a COMPLETE program rather than a PROCEDURE DIVISION
    // fragment. Wrapping one again nests it inside `PROGRAM-ID. T`, which cobc
    // rejects three different ways: "redefinition of program ID", "multiple
    // PROGRAM-ID's without matching END PROGRAM", and "CONFIGURATION SECTION
    // not allowed in nested programs".
    if body.to_ascii_uppercase().contains("IDENTIFICATION DIVISION") {
        // A COMPLETE program is not re-wrapped, so the scratch field and the
        // expected table have to be injected into ITS working storage —
        // returning the body untouched left `WS-VYBE-E` undeclared and every
        // runtime-paired case failed to compile.
        let mut decls = String::new();
        if needs_scratch {
            decls.push_str("\n01 WS-VYBE-L PIC X(256).");
        }
        if !extra_data.is_empty() {
            decls.push('\n');
            decls.push_str(extra_data);
        }
        if decls.is_empty() {
            return format!("{}\n", body.trim_end());
        }
        let upper = body.to_ascii_uppercase();
        let Some(at) = upper.find("WORKING-STORAGE SECTION.") else {
            return format!("{}\n", body.trim_end());
        };
        let at = at + "WORKING-STORAGE SECTION.".len();
        return format!("{}{decls}{}\n", &body[..at], body[at..].trim_end());
    }
    let mut data = prelude.map(str::trim).unwrap_or("").to_string();
    if needs_scratch {
        // The field the checks concatenate into. 256 is wider than any
        // expected line in the corpus; a shorter field would truncate and the
        // comparison would fail for the wrong reason.
        if !data.is_empty() {
            data.push('\n');
        }
        data.push_str("01 WS-VYBE-L PIC X(256).");
    }
    if !extra_data.is_empty() {
        if !data.is_empty() {
            data.push('\n');
        }
        data.push_str(extra_data);
    }
    let data = data.as_str();
    // Byte-for-byte the layout the test's own `p(data, body)` produced —
    // unindented division headers, four-space `STOP RUN.`
    format!(
        "IDENTIFICATION DIVISION.\nPROGRAM-ID. T.\nDATA DIVISION.\nWORKING-STORAGE SECTION.\n\
         {data}\nPROCEDURE DIVISION.\n{}\n    STOP RUN.\n",
        body.trim_end().trim_end_matches("STOP RUN.").trim_end()
    )
}

/// (statement start, insert point after the `.`, the operand list)
type Display = (usize, usize, Vec<String>);

/// Every `DISPLAY <operands>.` — one operand, several, or a bare literal. The
/// check concatenates the whole list, so no shape is excluded.
fn find_displays(src: &str) -> Vec<Display> {
    let mut out = Vec::new();
    let mut at = 0usize;
    while let Some(found) = src[at..].find("DISPLAY ") {
        let start = at + found;
        let before = src[..start].chars().last().unwrap_or('\n');
        let Some(dot) = statement_end(src, start) else {
            break;
        };
        at = dot + 1;
        // Not `DISPLAY` inside a literal, and not the FAIL line of a check we
        // just emitted.
        if before == '"' || before == '\'' {
            continue;
        }
        let text = src[start + "DISPLAY ".len()..dot].trim();
        let upper = text.to_ascii_uppercase();
        // Trailing clauses change what reaches stdout, so the concatenation
        // would not be the line.
        if upper.contains(" UPON ") || upper.contains("NO ADVANCING") {
            continue;
        }
        let operands = split_operands(text);
        if operands.is_empty() {
            continue;
        }
        out.push((start, dot + 1, operands));
    }
    out
}

/// The `.` that ends the statement, skipping any inside a literal.
fn statement_end(src: &str, from: usize) -> Option<usize> {
    let mut quote: Option<char> = None;
    for (i, ch) in src[from..].char_indices() {
        match quote {
            Some(q) if ch == q => quote = None,
            Some(_) => {}
            None => match ch {
                '"' | '\'' => quote = Some(ch),
                '.' => return Some(from + i),
                _ => {}
            } }
    }
    None
}

fn unpairable(src: &str, displays: &[Display], expected: &[String]) -> Option<String> {
    if displays.is_empty() {
        return Some("no DISPLAY to pair".into());
    }
    if src.contains("PERFORM") && (src.contains("TIMES") || src.contains("UNTIL")) {
        return Some("loop — DISPLAY count is not static".into());
    }
    // A DISPLAY inside a conditional clause is not a complete statement, so a
    // check appended after its `.` lands in the middle of the enclosing verb.
    for clause in ["ON SIZE ERROR", "AT END", "INVALID KEY", "ON EXCEPTION", "ON OVERFLOW"] {
        if src.contains(clause) {
            return Some(format!("DISPLAY inside an `{clause}` clause — cannot append a check"));
        }
    }
    if displays.len() != expected.len() {
        return Some(format!(
            "{} DISPLAY(s) but {} expected line(s)",
            displays.len(),
            expected.len()
        ));
    }
    None
}
