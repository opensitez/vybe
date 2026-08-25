//! Swapping the browser engine at RUN time.
//!
//! Its own test binary because the choice is process-global — one engine is
//! live per process, exactly as in a real run. A test in a shared binary would
//! be racing every other test's `install()`.
//!
//!     cargo test -p vybe_platform_web --features engine-htmlbox --test engine_toggle

use vybe_platform_web::engine_select::{self, Engine};

#[test]
fn the_engine_is_chosen_at_run_time_and_reports_what_it_installed() {
    // Nothing is live until something installs.
    assert_eq!(engine_select::live(), None, "an engine installed itself");

    let have = engine_select::available();
    assert!(
        have.contains(&Engine::Widgets),
        "a `gui` build always has the toolkit engine"
    );

    // A build with only one engine can only answer with that one; a build with
    // both must honour the choice. Both are asserted, so this test says
    // something true under either feature set.
    #[cfg(feature = "engine-htmlbox")]
    {
        assert!(
            have.contains(&Engine::HtmlBox),
            "`engine-htmlbox` is on but htmlbox is not available"
        );

        engine_select::choose(Engine::HtmlBox);
        assert_eq!(engine_select::install(), Some(Engine::HtmlBox));
        assert_eq!(engine_select::live(), Some(Engine::HtmlBox));

        // And back again, in the same process — the point of a runtime toggle.
        engine_select::choose(Engine::Widgets);
        assert_eq!(engine_select::install(), Some(Engine::Widgets));
        assert_eq!(engine_select::live(), Some(Engine::Widgets));
    }

    #[cfg(not(feature = "engine-htmlbox"))]
    {
        assert!(
            !have.contains(&Engine::HtmlBox),
            "htmlbox reported available in a build that did not compile it"
        );
        // Asking for an engine this build does not have must not silently
        // succeed with the other one pretending to be it: `install` answers
        // what it ACTUALLY installed, which is the toolkit.
        engine_select::choose(Engine::HtmlBox);
        assert_eq!(engine_select::install(), Some(Engine::Widgets));
        assert_eq!(engine_select::live(), Some(Engine::Widgets));
    }
}

#[test]
fn an_engine_name_is_parsed_the_way_a_user_would_type_it() {
    assert_eq!(Engine::parse("htmlbox"), Some(Engine::HtmlBox));
    assert_eq!(Engine::parse("HTMLBOX"), Some(Engine::HtmlBox));
    assert_eq!(Engine::parse(" widgets "), Some(Engine::Widgets));
    // The crate names work too, because that is what someone reading the
    // source would reach for.
    assert_eq!(Engine::parse("rhtmledit"), Some(Engine::HtmlBox));
    assert_eq!(Engine::parse("vybe_widgets"), Some(Engine::Widgets));
    // An unknown name is ignored rather than fatal — `VYBE_ENGINE=chrome`
    // falls back to the default instead of refusing to start.
    assert_eq!(Engine::parse("chrome"), None);
}
