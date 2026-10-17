use super::Utf16View;
use std::sync::Arc;

#[test]
fn ascii_selection_never_materializes_utf16_buffers() {
    let roots: [Arc<str>;3] = [Arc::from("abc\0".repeat(16384)),Arc::from("xyz"),Arc::from("")];
    let mut view=Utf16View::default();
    for _ in 0..10 {
        for root in &roots {
            view.select(root);
            assert!(view.ascii);
            assert!(view.units.is_empty());
            assert_eq!(view.units.capacity(),0);
            assert!(view.previous.iter().all(|entry|entry.units.capacity()==0));
            for (i,byte) in root.bytes().enumerate() {
                assert_eq!(byte as u16,root.as_bytes()[i] as u16);
            }
        }
    }
}

#[test]
fn lazy_ascii_and_decoded_unicode_views_keep_identity_and_exact_units() {
    let roots: [Arc<str>;3] = [Arc::from("ASCII\0"),Arc::from("\u{03a9}\u{1f980}"),Arc::from("xyz")];
    let mut view=Utf16View::default();
    for round in 0..30 {
        for root in &roots {
            view.select(root);
            assert_eq!(view.ascii,root.is_ascii());
            if !view.ascii {assert_eq!(view.units,root.encode_utf16().collect::<Vec<_>>())}
            if round%2==0 {assert_eq!(view.units(root),root.encode_utf16().collect::<Vec<_>>())}
            assert!(Arc::ptr_eq(view.source.as_ref().unwrap(),root));
        }
    }
    let dead: Arc<str> = Arc::from("dead".repeat(16384));
    let weak=Arc::downgrade(&dead);
    view.select(&dead);
    drop(dead);
    assert!(weak.upgrade().is_some());
    view.select(&roots[0]);
    assert!(weak.upgrade().is_none());
}
