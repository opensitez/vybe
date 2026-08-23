//! Which browser engine is live, chosen at RUN time.
//!
//! `engine::set_engine` and `canvas_backend::set_backend` are runtime slots, so
//! the choice never had to be a build-time one — both engines can be linked and
//! either can be installed. What made it look like a compile-time switch was
//! `register()` installing them in a fixed order, so the last `#[cfg]` to fire
//! always won.
//!
//! **The two have to move together.** An engine and a canvas painter that
//! disagree is not half a swap, it is a broken one: paint ops go to a document
//! that does not contain the node they name, so every drawing lands nowhere
//! while every call succeeds. [`install`] is the only place either is chosen,
//! which is what keeps them in step.

use std::sync::RwLock;

/// A browser engine this build can install.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Engine {
    /// `vybe_widgets` — the toolkit's own layout and painting.
    Widgets,
    /// `htmlbox` (`rhtmledit`) — the HTML/CSS engine.
    HtmlBox,
}

impl Engine {
    pub fn as_str(&self) -> &'static str {
        match self {
            Engine::Widgets => "widgets",
            Engine::HtmlBox => "htmlbox",
        }
    }

    /// Parse an engine name, case-insensitively. `None` for anything else.
    pub fn parse(name: &str) -> Option<Engine> {
        match name.trim().to_ascii_lowercase().as_str() {
            "widgets" | "vybe_widgets" => Some(Engine::Widgets),
            "htmlbox" | "rhtmledit" => Some(Engine::HtmlBox),
            _ => None,
        }
    }

    /// Whether this build has the engine compiled in at all.
    ///
    /// A choice this answers `false` for cannot be honoured, and [`install`]
    /// says so on stderr rather than quietly using the other one — a silent
    /// fallback would have someone comparing two engines and unknowingly
    /// measuring the same one twice.
    pub fn is_available(&self) -> bool {
        match self {
            Engine::Widgets => cfg!(feature = "gui"),
            Engine::HtmlBox => cfg!(feature = "engine-htmlbox"),
        }
    }
}

/// Every engine this build can install.
pub fn available() -> Vec<Engine> {
    [Engine::Widgets, Engine::HtmlBox]
        .into_iter()
        .filter(Engine::is_available)
        .collect()
}

/// The default when nothing asks for anything.
///
/// `vybe_widgets`, because it is what every existing caller has been getting —
/// .NET's designer, Flutter's realizer and SDL all reach the toolkit today, and
/// a default that changed under them would be a swap nobody asked for.
const DEFAULT: Engine = Engine::Widgets;

/// An explicit choice made in-process, which beats the environment.
static CHOICE: RwLock<Option<Engine>> = RwLock::new(None);

/// The engine currently installed, once [`install`] has run.
static LIVE: RwLock<Option<Engine>> = RwLock::new(None);

/// Choose the engine for this process, before `register()` runs.
///
/// Beats `VYBE_ENGINE`, so a host that knows what it wants is not overridden by
/// an environment variable someone left set.
pub fn choose(engine: Engine) {
    *CHOICE.write().unwrap() = Some(engine);
}

/// Which engine is live. `None` before `register()` has installed one.
pub fn live() -> Option<Engine> {
    *LIVE.read().unwrap()
}

/// What was asked for, before availability is considered.
///
/// An in-process [`choose`] first, then `VYBE_ENGINE`, then the default.
///
/// A `VYBE_ENGINE` value that does not name an engine is not fatal — a run
/// should not refuse to start over an environment variable — but it SAYS so.
/// Falling back in silence is how `VYBE_ENGINE=htmlbx` gets you the toolkit
/// engine and a comparison that measures it against itself.
fn requested() -> Engine {
    if let Some(chosen) = *CHOICE.read().unwrap() {
        return chosen;
    }
    let Ok(name) = std::env::var("VYBE_ENGINE") else {
        return DEFAULT;
    };
    match Engine::parse(&name) {
        Some(engine) => engine,
        None => {
            eprintln!(
                "vybe: VYBE_ENGINE=`{name}` names no engine (try: {}); using `{}`.",
                available()
                    .iter()
                    .map(|e| e.as_str())
                    .collect::<Vec<_>>()
                    .join(", "),
                DEFAULT.as_str(),
            );
            DEFAULT
        }
    }
}

/// Install the chosen engine and its canvas painter.
///
/// Both, together, and nothing else installs either — see the module note on
/// why an engine without its matching painter is worse than no swap at all.
///
/// Answers what it actually installed, which is not always what was asked for:
/// a build without `engine-htmlbox` has no htmlbox to install.
pub fn install() -> Option<Engine> {
    let want = requested();
    let engine = if want.is_available() {
        want
    } else {
        // Loud, because the alternative is someone comparing two engines and
        // measuring the same one twice without knowing it.
        let have = available();
        eprintln!(
            "vybe: engine `{}` is not in this build (have: {}); using `{}`. \
             Rebuild with `--features vybe_platform_web/engine-htmlbox` for htmlbox.",
            want.as_str(),
            if have.is_empty() {
                "none".to_string()
            } else {
                have.iter()
                    .map(|e| e.as_str())
                    .collect::<Vec<_>>()
                    .join(", ")
            },
            have.first().map(Engine::as_str).unwrap_or("none"),
        );
        *available().first()?
    };

    match engine {
        Engine::Widgets => {
            #[cfg(feature = "gui")]
            {
                crate::engine_widgets::install();
                crate::canvas_backend_widgets::install();
            }
        }
        Engine::HtmlBox => {
            #[cfg(feature = "engine-htmlbox")]
            {
                crate::engine_htmlbox::install();
                crate::canvas_backend_htmlbox::install();
            }
        }
    }
    *LIVE.write().unwrap() = Some(engine);
    Some(engine)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn an_engine_name_round_trips() {
        for e in [Engine::Widgets, Engine::HtmlBox] {
            assert_eq!(Engine::parse(e.as_str()), Some(e));
        }
        assert_eq!(Engine::parse("  HtmlBox "), Some(Engine::HtmlBox));
        assert_eq!(Engine::parse("chrome"), None, "an engine we do not have");
    }

    #[test]
    fn the_build_reports_what_it_actually_contains() {
        // `available()` is derived from the features, so it cannot claim an
        // engine this build did not compile.
        let have = available();
        assert_eq!(
            have.contains(&Engine::HtmlBox),
            cfg!(feature = "engine-htmlbox")
        );
        assert_eq!(have.contains(&Engine::Widgets), cfg!(feature = "gui"));
    }
}
