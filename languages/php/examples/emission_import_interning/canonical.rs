//! Canonicalize only unordered brace groups. Square/round delimiters and their
//! entries remain ordered; quoted strings (including escaped delimiters) stay
//! byte-for-byte intact.
pub fn canonical(text: &str) -> String {
    fn group(input: &[u8], i: &mut usize, close: Option<u8>, sort: bool) -> Vec<u8> {
        let mut output = Vec::new();
        let mut fields = Vec::new();
        while *i < input.len() {
            let b = input[*i];
            if Some(b) == close { *i += 1; break; }
            match b {
                b'"' => {
                    let start = *i;
                    *i += 1;
                    let mut escaped = false;
                    while *i < input.len() {
                        let c = input[*i]; *i += 1;
                        if escaped { escaped = false; }
                        else if c == b'\\' { escaped = true; }
                        else if c == b'"' { break; }
                    }
                    output.extend_from_slice(&input[start..*i]);
                }
                b'{' | b'[' | b'(' => {
                    let end = match b { b'{' => b'}', b'[' => b']', _ => b')' };
                    *i += 1;
                    output.push(b);
                    output.extend(group(input, i, Some(end), b == b'{'));
                    output.push(end);
                }
                b',' if sort => {
                    fields.push(std::mem::take(&mut output)); *i += 1;
                }
                b'}' | b']' | b')' => panic!("mismatched Debug delimiter"),
                _ => { output.push(b); *i += 1; }
            }
        }
        if sort {
            fields.push(output);
            let mut fields: Vec<_> = fields.into_iter().map(|field|
                String::from_utf8(field).unwrap().trim().to_owned()).collect();
            fields.sort(); fields.join(", ").into_bytes()
        } else { output }
    }
    String::from_utf8(group(text.as_bytes(), &mut 0, None, false)).unwrap()
}

#[cfg(test)]
mod tests {
    use super::canonical;
    #[test] fn ordered_vectors_and_tuples_are_never_sorted() {
        assert_ne!(canonical("Chunk { code: [1, 2], data: (3, 4) }"), canonical("Chunk { code: [2, 1], data: (3, 4) }"));
        assert_ne!(canonical("Chunk { code: [1, 2], data: (3, 4) }"), canonical("Chunk { code: [1, 2], data: (4, 3) }"));
        assert_ne!(canonical("[{ a: 1 }, { a: 2 }]"), canonical("[{ a: 2 }, { a: 1 }]"));
    }
    #[test] fn nested_map_and_named_field_order_are_canonical() {
        assert_eq!(canonical("Chunk { x: [{ b: 2, a: 1 }], y: (1, { c: 3, d: 4 }) }"), canonical("Chunk { y: (1, { d: 4, c: 3 }), x: [{ a: 1, b: 2 }] }"));
    }
    #[test] fn quoted_escapes_unicode_and_delimiters_are_exact() {
        let a = r#"Value { z: "é, [2,1] {b,a} (4,3) \" \\", a: [2, 1] }"#;
        let b = r#"Value { a: [2, 1], z: "é, [2,1] {b,a} (4,3) \" \\" }"#;
        assert_eq!(canonical(a), canonical(b));
        assert_ne!(canonical(r#"{ s: "[1,2]" }"#), canonical(r#"{ s: "[2,1]" }"#));
    }
}
