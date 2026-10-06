use vybe_ast::{RecordByteField, RecordByteLayout};

fn field(name: &str, offset: u64, bytes: u64) -> RecordByteField {
    RecordByteField {
        name: name.into(),
        offset,
        bytes,
    }
}

#[test]
fn padding_and_overlapping_members_keep_declared_coordinates() {
    let layout = RecordByteLayout {
        bytes: 16,
        align: 8,
        fields: vec![
            field("tag", 0, 1),
            field("integer", 8, 8),
            field("real", 8, 8),
        ],
    };
    layout.validate().unwrap();
    assert_eq!(layout.field("integer").unwrap().offset, 8);
    assert!(layout.field("padding").is_none());
}

#[test]
fn invalid_extents_alignment_and_duplicate_names_are_rejected() {
    for layout in [
        RecordByteLayout {
            bytes: 8,
            align: 0,
            fields: vec![],
        },
        RecordByteLayout {
            bytes: 8,
            align: 3,
            fields: vec![],
        },
        RecordByteLayout {
            bytes: 7,
            align: 4,
            fields: vec![],
        },
        RecordByteLayout {
            bytes: 8,
            align: 4,
            fields: vec![field("x", 7, 4)],
        },
        RecordByteLayout {
            bytes: 8,
            align: 4,
            fields: vec![field("x", u64::MAX, 4)],
        },
        RecordByteLayout {
            bytes: 8,
            align: 4,
            fields: vec![field("x", 0, 4), field("x", 4, 4)],
        },
    ] {
        assert!(layout.validate().is_err(), "accepted {layout:?}");
    }
}
