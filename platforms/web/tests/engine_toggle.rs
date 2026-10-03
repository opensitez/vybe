//! Swapping the browser engine at RUN time.
//!
//! Its own test binary because the choice is process-global — one engine is
//! live per process, exactly as in a real run. A test in a shared binary would
//! be racing every other test's `install()`.
//!
//!     cargo test -p vybe_platform_web --features engine-webcore --test engine_toggle

use vybe_platform_web::engine_select::{self, Engine};

#[test]
fn the_engine_is_chosen_at_run_time_and_reports_what_it_installed() {
    // Nothing is live until something installs.
    assert_eq!(engine_select::live(), None, "an engine installed itself");

    let have = engine_select::available();
    assert!(!have.is_empty(), "the browser build has no engine");

    assert!(have.contains(&Engine::WebCore));
    engine_select::choose(Engine::WebCore);
    assert_eq!(engine_select::install(), Some(Engine::WebCore));
    assert_eq!(engine_select::live(), Some(Engine::WebCore));
}

#[test]
fn an_engine_name_is_parsed_the_way_a_user_would_type_it() {
    assert_eq!(Engine::parse("webcore"), Some(Engine::WebCore));
    assert_eq!(Engine::parse("WEBCORE"), Some(Engine::WebCore));
    assert_eq!(Engine::parse(" osbrowser "), Some(Engine::OsBrowser));
    assert_eq!(Engine::parse("widgets"), None);
    // An unknown name is ignored rather than fatal — `VYBE_ENGINE=chrome`
    // falls back to the default instead of refusing to start.
    assert_eq!(Engine::parse("chrome"), None);
}
