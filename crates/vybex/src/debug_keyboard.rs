//! Keyboard command parsing, independent of the VM and browser engines.

pub(crate) struct KeyboardInput {
    pub kind: String,
    pub key: String,
    pub key_code: i32,
    pub code: String,
}

pub(crate) fn parse(parts: &[&str]) -> Result<KeyboardInput, String> {
    let usage = "usage: key <keydown|keyup> <key> <keyCode> [code]";
    if !(3..=4).contains(&parts.len()) || !matches!(parts[0], "keydown" | "keyup") {
        return Err(usage.into());
    }
    let key_code = parts[2].parse::<i32>().map_err(|_| "keyCode must be a nonnegative i32")?;
    if key_code < 0 {
        return Err("keyCode must be a nonnegative i32".into());
    }
    let key = if parts[1].starts_with('"') {
        serde_json::from_str::<String>(parts[1])
            .map_err(|_| "key must be a valid JSON string when quoted")?
    } else {
        parts[1].into()
    };
    Ok(KeyboardInput {
        kind: parts[0].into(), key, key_code,
        code: parts.get(3).copied().unwrap_or("").into(),
    })
}

#[cfg(test)]
mod tests {
    use super::parse;

    #[test]
    fn preserves_standard_event_fields() {
        let event = parse(&["keydown", "ArrowUp", "38", "ArrowUp"]).unwrap();
        assert_eq!(event.kind, "keydown");
        assert_eq!(event.key, "ArrowUp");
        assert_eq!(event.key_code, 38);
        assert_eq!(event.code, "ArrowUp");
        let release = parse(&["keyup", "a", "65"]).unwrap();
        assert_eq!(release.kind, "keyup");
        assert_eq!(release.key, "a");
        assert_eq!(release.key_code, 65);
        assert_eq!(release.code, "");
    }

    #[test]
    fn rejects_invalid_or_extra_arguments() {
        for parts in [vec![], vec!["keydown"], vec!["click", "a", "65"],
            vec!["keydown", "a", "-1"], vec!["keyup", "a", "2147483648"],
            vec!["keydown", "a", "65", "KeyA", "ignored"]] {
            assert!(parse(&parts).is_err(), "{parts:?}");
        }
    }

    #[test]
    fn decodes_json_quoted_key_tokens() {
        for kind in ["keydown", "keyup"] {
            let command = format!(r#"{kind} "\u0020" 32 Space"#);
            let parts = command.split_whitespace().collect::<Vec<_>>();
            let event = parse(&parts).unwrap();
            assert_eq!(event.kind, kind);
            assert_eq!(event.key, " ");
            assert_eq!(event.key_code, 32);
            assert_eq!(event.code, "Space");
        }
        assert_eq!(parse(&["keydown", r#""""#, "0"]).unwrap().key, "");
        assert_eq!(parse(&["keydown", r#""ArrowUp""#, "38", "ArrowUp"]).unwrap().key, "ArrowUp");
        assert_eq!(parse(&["keydown", "Space", "32", "Space"]).unwrap().key, "Space");
    }

    #[test]
    fn rejects_malformed_json_quoted_keys() {
        for key in ["\"", "\"Space", r#""\q""#, r#""\u002""#, r#""a"tail"#] {
            assert!(parse(&["keydown", key, "32", "Space"]).is_err(), "{key:?}");
        }
    }
}
