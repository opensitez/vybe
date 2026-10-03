use std::sync::Arc;
use vybe_parser::{
    MatchOptions,
    editor::{EditError, Report, Session, TextRange},
    source::{Encoding, Position, PositionError},
};
fn position(character: u32) -> Position {
    Position { line: 0, character }
}
fn session(source: &str) -> Session {
    let grammar =
        Arc::new(vybe_parser::compile("root = { SOI ~ cell* ~ EOI } cell = { ANY }").unwrap());
    Session::new(grammar, "root", source.into(), MatchOptions::default())
}

#[test]
fn utf16_edits_are_atomic_and_old_reports_keep_source_and_node_identity() {
    let mut session = session("a😀b");
    let old = session.parse();
    let old_id = old.captures().next().unwrap().id;
    assert_eq!(old.text(old_id), Some("a😀b"));
    let range = TextRange {
        start: position(1),
        end: position(3),
    };
    assert_eq!(session.edit(0, range, Encoding::Utf16, "é\n"), Ok(1));
    let new = session.parse();
    assert_eq!(session.source(), "aé\nb");
    assert_eq!(new.revision(), 1);
    assert_eq!(old.source(), "a😀b");
    assert_eq!(old.text(old_id), Some("a😀b"));
    assert_eq!(new.text(old_id), None);
    assert_eq!(
        new.index().position(4, Encoding::Utf16),
        Ok(Position {
            line: 1,
            character: 0
        })
    );
    assert_eq!(
        session.edit(0, range, Encoding::Utf16, "x"),
        Err(EditError::StaleRevision)
    );
    assert_eq!(session.source(), "aé\nb");
    fn shareable<T: Send + Sync>() {}
    shareable::<Report>();
}

#[test]
fn invalid_scalar_ranges_and_reversed_edits_do_not_change_source_or_revision() {
    let mut session = session("😀z");
    for (start, end, error) in [
        (1, 2, EditError::Position(PositionError::InsideScalar)),
        (2, 1, EditError::Position(PositionError::InsideScalar)),
        (3, 2, EditError::ReversedRange),
        (0, 10, EditError::Position(PositionError::OutOfBounds)),
    ] {
        assert_eq!(
            session.edit(
                0,
                TextRange {
                    start: position(start),
                    end: position(end)
                },
                Encoding::Utf16,
                "x"
            ),
            Err(error)
        );
        assert_eq!(session.revision(), 0);
        assert_eq!(session.source(), "😀z");
    }
    let a = session.parse();
    let b = self::session("other document").parse();
    assert_eq!(b.text(a.captures().next().unwrap().id), None);
}

#[test]
fn reports_own_correct_revision_diagnostics_and_require_complete_documents() {
    let grammar = Arc::new(vybe_parser::compile("number = { ASCII_DIGIT+ }").unwrap());
    let mut session = Session::new(grammar, "number", "12".into(), MatchOptions::default());
    let valid = session.parse();
    assert!(valid.error().is_none());
    session
        .edit(
            0,
            TextRange {
                start: position(2),
                end: position(2),
            },
            Encoding::Utf16,
            "x",
        )
        .unwrap();
    let invalid = session.parse();
    assert_eq!(invalid.revision(), 1);
    assert_eq!(invalid.error_position(Encoding::Utf16), Some(position(2)));
    assert_eq!(invalid.captures().count(), 0);
    assert!(invalid.consumed().is_none());
    assert_eq!(valid.source(), "12");
    assert_eq!(invalid.source(), "12x");
}
