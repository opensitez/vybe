use vybe_parser::source::{Encoding, Index, Position, PositionError};

#[test]
fn every_representable_scalar_boundary_round_trips_all_encodings() {
    let source = format!("aé😀\r\n{}\rfinal\n", "α😀éabc".repeat(300));
    let index = Index::new(&source);
    assert_eq!(index.line_count(), 4);
    for byte in 0..=source.len() {
        if !source.is_char_boundary(byte) {
            continue;
        }
        for encoding in [Encoding::Utf8, Encoding::Utf16, Encoding::Scalar] {
            match index.position(byte, encoding) {
                Ok(position) => assert_eq!(index.offset(position, encoding).unwrap(), byte),
                Err(PositionError::InsideLineBreak) => {
                    assert_eq!(source.as_bytes()[byte - 1..=byte], *b"\r\n")
                }
                error => panic!("{byte}: {error:?}"),
            }
        }
    }
}

#[test]
fn utf16_surrogates_utf8_boundaries_crlf_and_bounds_are_explicit_errors() {
    let index = Index::new("😀é\r\nz");
    assert_eq!(
        index.offset(
            Position {
                line: 0,
                character: 1
            },
            Encoding::Utf16
        ),
        Err(PositionError::InsideScalar)
    );
    assert_eq!(
        index.offset(
            Position {
                line: 0,
                character: 2
            },
            Encoding::Utf16
        ),
        Ok(4)
    );
    assert_eq!(
        index.offset(
            Position {
                line: 0,
                character: 1
            },
            Encoding::Utf8
        ),
        Err(PositionError::InsideScalar)
    );
    assert_eq!(
        index.position(1, Encoding::Utf16),
        Err(PositionError::InsideScalar)
    );
    assert_eq!(
        index.position(7, Encoding::Utf16),
        Err(PositionError::InsideLineBreak)
    );
    assert_eq!(
        index.position(8, Encoding::Utf16),
        Ok(Position {
            line: 1,
            character: 0
        })
    );
    assert_eq!(
        index.offset(
            Position {
                line: 2,
                character: 0
            },
            Encoding::Utf16
        ),
        Err(PositionError::OutOfBounds)
    );
    assert_eq!(
        index.offset(
            Position {
                line: 0,
                character: 99
            },
            Encoding::Utf16
        ),
        Err(PositionError::OutOfBounds)
    );
    assert_eq!(
        index.position(100, Encoding::Utf16),
        Err(PositionError::OutOfBounds)
    );
    assert_eq!(
        Index::new("").position(0, Encoding::Utf16),
        Ok(Position {
            line: 0,
            character: 0
        })
    );
}
