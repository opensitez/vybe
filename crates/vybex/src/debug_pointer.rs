pub(super) struct PointerInput {
    pub kind: String,
    pub x: f32,
    pub y: f32,
    pub button: i32,
}

pub(super) fn parse(parts: &[&str]) -> Result<PointerInput, &'static str> {
    if !(3..=4).contains(&parts.len()) {
        return Err("usage: pointer <mousedown|mouseup|mousemove> <x> <y> [button]");
    }
    if !matches!(parts[0], "mousedown" | "mouseup" | "mousemove") {
        return Err("pointer kind must be mousedown, mouseup, or mousemove");
    }
    let coordinate = |text: &str| {
        text.parse::<f32>().ok().filter(|value| value.is_finite())
            .ok_or("pointer coordinates must be finite numbers")
    };
    let button = parts.get(3).unwrap_or(&"0").parse::<i32>()
        .map_err(|_| "button must be an integer from 0 to 4")?;
    if !(0..=4).contains(&button) {
        return Err("button must be an integer from 0 to 4");
    }
    Ok(PointerInput {
        kind: parts[0].into(), x: coordinate(parts[1])?, y: coordinate(parts[2])?, button,
    })
}

#[cfg(test)]
mod tests {
    use super::parse;

    #[test]
    fn pointer_arguments_preserve_client_coordinates_and_button() {
        let event = parse(&["mousedown", "100.5", "-20", "2"]).unwrap();
        assert_eq!(event.kind, "mousedown");
        assert_eq!((event.x, event.y, event.button), (100.5, -20.0, 2));
        assert_eq!(parse(&["mouseup", "0", "0"]).unwrap().button, 0);
        for parts in [vec![], vec!["click", "0", "0"], vec!["mousemove", "NaN", "0"],
            vec!["mousedown", "0", "inf"], vec!["mouseup", "0", "0", "-1"],
            vec!["mouseup", "0", "0", "5"], vec!["mouseup", "0", "0", "0", "extra"]] {
            assert!(parse(&parts).is_err(), "{parts:?}");
        }
    }
}
