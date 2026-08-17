//! **Selectors** — the Selectors Level 4 subset, as a parser and a matcher.
//!
//! This is the input the cascade never had. `vybe_widgets` has exactly two
//! origins — a UA table keyed on a bare tag name, and the inline `style=""`
//! store — not because a third was rejected but because nothing could express
//! one: a rule needs a selector, and there was no selector language here.
//!
//! ## Why it lives in the toolkit
//!
//! The engine already existed, in `platforms/web/src/dom_parser.rs`, wired to
//! `querySelector`/`querySelectorAll` on a DOM that **renders nothing** (see
//! the two-DOMs problem). `platforms/web` depends on `vybe_widgets` and not the
//! other way round, so the parser could only be shared by moving it DOWN here.
//! Writing a second one would have been the wrong instinct twice over: two
//! engines to disagree, and one of them in the layer that cannot use it.
//!
//! Parsing and matching are split deliberately. A selector is pure syntax and
//! belongs to neither tree; matching needs a node, and the two DOMs have
//! different ones. So the parser is shared outright and each tree brings its
//! own [`Element`].
//!
//! ## Pseudo-classes: exactly the ones that can be ANSWERED
//!
//! `:hover` and `:checked` match, because the widget knows both — the panel
//! that arranges the controls maintains `hovered()`, and `checked` has a
//! command. [`Element::state`] is how the tree answers, and the live document
//! serves it from a cache it refreshes on a sweep, so a state change also
//! re-runs that element's cascade.
//!
//! Everything else is still refused WHOLE. `:focus` and `:disabled` read as
//! though they were equally available and are not: the form tracks focus by
//! child index rather than by node, and `SetEnabled` has no `Get` counterpart.
//! `::before` needs a generated box; `:nth-child()` needs sibling
//! invalidation. A selector that parses and then silently never matches is
//! worse than one that is refused, because it looks supported — so an
//! unanswerable pseudo-class makes the whole selector `None` rather than a
//! selector that quietly drops a condition.

/// One simple selector — the atoms a compound is built from.
#[derive(Debug, Clone, PartialEq)]
pub enum SimplePart {
    Universal,
    Type(String),
    Id(String),
    Class(String),
    Attr {
        name: String,
        op: AttrOp,
        value: Option<String>,
    },
    /// A pseudo-class that asks about the element's live STATE — `:hover`,
    /// `:focus`, `:active`, `:disabled`, `:enabled`, `:checked`.
    ///
    /// A condition on this element, not a generated box, which is why it is a
    /// simple part like the rest. What it cannot be is derived from the tree:
    /// only the widget knows whether the pointer is over it, so the answer
    /// comes back through [`Element::state`].
    State(String),
    /// A pseudo-class about the element's POSITION among its siblings —
    /// `:first-child`, `:last-child`, `:only-child`, `:nth-child()`,
    /// `:nth-last-child()`.
    ///
    /// Every one of these is `an+b` underneath, which is why they share a
    /// variant instead of getting five: `:first-child` is `:nth-child(0n+1)`
    /// and `:last-child` is the same counted from the end. Collapsing them at
    /// the PARSER means one matcher, and no chance of the shorthands drifting
    /// from the general form.
    NthChild {
        /// The `a` of `an+b` — the period. Zero means a fixed position.
        step: i32,
        /// The `b` of `an+b` — the offset, 1-based like CSS.
        offset: i32,
        /// Counted from the END, for `:last-child`/`:nth-last-child()`.
        from_end: bool,
        /// Count only siblings with the SAME TAG — the `-of-type` family.
        of_type: bool,
    },
    /// `:not(…)` — the negation. Holds a full selector LIST, because
    /// `:not(a, .b)` is "neither", and each alternative is a compound.
    ///
    /// Recursive by construction: the inner selectors are matched by the same
    /// matcher, so `:not(:first-child)` costs nothing extra to support.
    Not(Vec<CompoundSelector>),
}

/// Parse the `an+b` of `:nth-child(…)` — Selectors §9.
///
/// Accepts the two keywords and the full microsyntax: `odd`, `even`, `3`,
/// `2n`, `2n+1`, `-n+3`, `+3n-2`. Whitespace around the sign is allowed
/// because the spec allows it (`2n + 1`), and a stylesheet in the wild uses it.
fn parse_nth(input: &str) -> Option<(i32, i32)> {
    let text: String = input
        .chars()
        .filter(|c| !c.is_whitespace())
        .collect::<String>()
        .to_ascii_lowercase();
    match text.as_str() {
        "odd" => return Some((2, 1)),
        "even" => return Some((2, 0)),
        "" => return None,
        _ => {}
    }
    let Some(n_at) = text.find('n') else {
        // No `n` at all — a fixed position, `:nth-child(3)`.
        return text.parse::<i32>().ok().map(|b| (0, b));
    };
    let (head, tail) = text.split_at(n_at);
    let step = match head {
        "" | "+" => 1,
        "-" => -1,
        other => other.parse::<i32>().ok()?,
    };
    let rest = &tail[1..];
    let offset = if rest.is_empty() {
        0
    } else {
        // `+3` parses; so must `-3`, which `i32::from_str` also accepts.
        rest.parse::<i32>().ok()?
    };
    Some((step, offset))
}

/// Does a 1-based position satisfy `an+b`?
///
/// The whole of the rule: some non-negative `n` must solve `position = a·n +
/// b`. `a == 0` is the fixed case, and a negative period counts DOWN from `b`,
/// which is what makes `:nth-child(-n+3)` mean "the first three".
fn nth_matches(position: i32, step: i32, offset: i32) -> bool {
    if step == 0 {
        return position == offset;
    }
    let delta = position - offset;
    // Same sign as the period and an exact multiple of it.
    delta % step == 0 && delta / step >= 0
}

/// The pseudo-classes this engine can answer. Everything else is refused at
/// parse time, so a rule never half-matches.
///
/// **Two, and only because the widget can be ASKED.** `hovered()` is on the
/// panel trait and the arranging panel maintains it; `checked` has a command.
///
/// `:focus`, `:active`, `:disabled` and `:enabled` are deliberately NOT here
/// even though they sound equally available: the form tracks focus by CHILD
/// INDEX rather than by node, and `SetEnabled` has no `Get` counterpart, so
/// nothing can answer them. Accepting them would produce a selector that parses
/// and then never matches — the exact failure the refusal above exists to
/// prevent, and worse than refusing because it looks supported.
///
/// The structural ones (`:first-child`, `:nth-child()`) are derivable from the
/// tree but need an invalidation story of their own: inserting a sibling
/// changes which elements match.
fn is_supported_state(name: &str) -> bool {
    matches!(name, "hover" | "checked")
}

/// How an attribute selector compares.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AttrOp {
    /// `[attr]` — present, whatever it holds.
    Has,
    /// `[attr="v"]`
    Exact,
    /// `[attr*="v"]`
    Substring,
    /// `[attr^="v"]`
    Prefix,
    /// `[attr$="v"]`
    Suffix,
    /// `[attr~="v"]` — one of a whitespace-separated list.
    Word,
    /// `[attr|="v"]` — exactly `v`, or `v-` prefixed. What `lang` is for.
    Lang,
}

/// Simple selectors with no combinator between them — `a.btn[disabled]`.
#[derive(Debug, Clone, PartialEq)]
pub struct CompoundSelector {
    pub parts: Vec<SimplePart>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Combinator {
    Descendant,
    Child,
    AdjacentSibling,
    GeneralSibling,
}

/// Compounds joined by combinators — `nav > ul li`.
#[derive(Debug, Clone, PartialEq)]
pub struct ComplexSelector {
    /// Pairs of (combinator-from-previous, compound). The first entry's
    /// combinator is unused (root of the chain).
    pub parts: Vec<(Combinator, CompoundSelector)>,
}

impl ComplexSelector {
    /// **Specificity** — the a-b-c triple, as one sortable number.
    ///
    /// The whole reason a stylesheet needs more than document order: `#main p`
    /// beats `p` however they are written, and `.warn` beats `div`. Counted
    /// over every compound in the chain, which is what makes a descendant
    /// selector more specific than the compound it ends with.
    ///
    /// Packed rather than compared field by field because the cascade sorts by
    /// it, and a single `u32` makes "specificity, then source order" an
    /// ordinary tuple sort. Each field is clamped at 255 — a selector with 256
    /// classes is not a real input, and saturating is what keeps the packing
    /// monotonic if one ever appears.
    /// Whether any part of this selector depends on the element's POSITION
    /// among its siblings.
    ///
    /// The cascade needs to know because those are the only matches a
    /// STRUCTURAL change can invalidate: appending a child makes the previous
    /// `:last-child` stop matching, and no restyle of the appended node would
    /// ever notice. A document whose stylesheet uses none of them pays nothing
    /// for the check, which is why it is asked of the sheet rather than
    /// assumed.
    pub fn is_positional(&self) -> bool {
        self.parts.iter().any(|(_, compound)| {
            compound
                .parts
                .iter()
                .any(|part| matches!(part, SimplePart::NthChild { .. }))
        })
    }

    pub fn specificity(&self) -> u32 {
        let (mut ids, mut classes, mut types) = (0u32, 0u32, 0u32);
        for (_, compound) in &self.parts {
            for part in &compound.parts {
                match part {
                    SimplePart::Id(_) => ids += 1,
                    // A pseudo-class weighs the same as a class — Selectors
                    // §17, which is what makes `a:hover` beat a bare `a`.
                    SimplePart::Class(_)
                    | SimplePart::Attr { .. }
                    | SimplePart::State(_)
                    | SimplePart::NthChild { .. } => classes += 1,
                    // `:not()` contributes NOTHING itself — its most specific
                    // argument does. Selectors §17.
                    SimplePart::Not(alternatives) => {
                        let inner = alternatives
                            .iter()
                            .map(|compound| {
                                ComplexSelector {
                                    parts: vec![(Combinator::Descendant, compound.clone())],
                                }
                                .specificity()
                            })
                            .max()
                            .unwrap_or(0);
                        ids += (inner >> 16) & 0xff;
                        classes += (inner >> 8) & 0xff;
                        types += inner & 0xff;
                    }
                    SimplePart::Type(_) => types += 1,
                    // The universal selector counts for nothing, by spec.
                    SimplePart::Universal => {}
                }
            }
        }
        (ids.min(255) << 16) | (classes.min(255) << 8) | types.min(255)
    }
}

/// What a tree must answer for a selector to be matched against it.
///
/// Small on purpose: everything the subset above needs and nothing else. A tree
/// that can answer these five questions can be selected over, which is what
/// lets one engine serve both the live widget DOM and the parsed one without
/// either knowing the other exists.
pub trait Element: Sized {
    fn tag(&self) -> String;
    fn attribute(&self, name: &str) -> Option<String>;
    fn parent(&self) -> Option<Self>;
    /// The element immediately before this one among its parent's children.
    fn previous_sibling(&self) -> Option<Self>;
    /// The element immediately after it. Needed by `:last-child` and
    /// `:nth-last-child()`, which cannot be answered by walking backwards.
    ///
    /// Defaulted to `None` so a tree that only ever looks backwards keeps
    /// compiling; such a tree simply never matches the forward-looking
    /// pseudo-classes, which is the honest answer for it.
    fn next_sibling(&self) -> Option<Self> {
        None
    }
    /// Whether a live state pseudo-class holds — `:hover`, `:focus`,
    /// `:active`, `:disabled`, `:enabled`, `:checked`.
    ///
    /// Defaulted to `false` because only a tree with WIDGETS can answer it. A
    /// property-bag document has no pointer and no controls, so every state is
    /// absent there — which is the right answer for it, not a stub.
    fn state(&self, _name: &str) -> bool {
        false
    }
}

/// Parse a comma-separated selector list. `None` if any of it is unsupported.
pub fn parse_selector_list(input: &str) -> Option<Vec<ComplexSelector>> {
    let mut out = Vec::new();
    for piece in split_top_level(input) {
        let trimmed = piece.trim();
        if trimmed.is_empty() {
            return None;
        }
        out.push(parse_complex(trimmed)?);
    }
    if out.is_empty() { None } else { Some(out) }
}

/// Split a selector list on its TOP-LEVEL commas.
///
/// A plain `split(',')` tears `p:not(.a, #b)` in half: the comma inside the
/// parentheses belongs to the negation's own list, not to the outer one. The
/// same shape as any nested-argument syntax, and the same fix — count the
/// depth and only break at zero.
fn split_top_level(input: &str) -> Vec<&str> {
    let mut out = Vec::new();
    let mut depth = 0usize;
    let mut start = 0usize;
    for (i, c) in input.char_indices() {
        match c {
            '(' => depth += 1,
            ')' => depth = depth.saturating_sub(1),
            ',' if depth == 0 => {
                out.push(&input[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    out.push(&input[start..]);
    out
}

fn parse_complex(input: &str) -> Option<ComplexSelector> {
    let chars: Vec<char> = input.chars().collect();
    let mut i = 0;
    let mut parts: Vec<(Combinator, CompoundSelector)> = Vec::new();
    let mut next_combinator = Combinator::Descendant;
    let mut first = true;

    while i < chars.len() {
        // Skip whitespace; remember if we crossed any (descendant combinator).
        let pre_ws_start = i;
        while i < chars.len() && chars[i].is_whitespace() {
            i += 1;
        }
        let had_whitespace = i > pre_ws_start;
        if i >= chars.len() {
            break;
        }
        let explicit = match chars[i] {
            '>' => Some(Combinator::Child),
            '+' => Some(Combinator::AdjacentSibling),
            '~' => Some(Combinator::GeneralSibling),
            _ => None,
        };
        if let Some(comb) = explicit {
            next_combinator = comb;
            i += 1;
            while i < chars.len() && chars[i].is_whitespace() {
                i += 1;
            }
        } else if had_whitespace && !first {
            next_combinator = Combinator::Descendant;
        }

        let (compound, consumed) = parse_compound(&chars[i..])?;
        if consumed == 0 {
            return None;
        }
        i += consumed;

        if first {
            parts.push((Combinator::Descendant, compound));
            first = false;
        } else {
            parts.push((next_combinator, compound));
        }
    }

    if parts.is_empty() {
        None
    } else {
        Some(ComplexSelector { parts })
    }
}

fn parse_compound(chars: &[char]) -> Option<(CompoundSelector, usize)> {
    let mut parts: Vec<SimplePart> = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c.is_whitespace() || c == '>' || c == '+' || c == '~' || c == ',' {
            break;
        }
        if c == '*' {
            parts.push(SimplePart::Universal);
            i += 1;
        } else if c == '#' || c == '.' {
            let is_id = c == '#';
            i += 1;
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            if start == i {
                return None;
            }
            let name: String = chars[start..i].iter().collect();
            parts.push(if is_id {
                SimplePart::Id(name)
            } else {
                SimplePart::Class(name)
            });
        } else if c == '[' {
            i += 1;
            let (attr, consumed) = parse_attr(&chars[i..])?;
            i += consumed;
            parts.push(attr);
        } else if is_ident_start(c) {
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            parts.push(SimplePart::Type(chars[start..i].iter().collect()));
        } else if c == ':' {
            // A pseudo-class — but only one whose state this tree can actually
            // answer. The refusal below still stands for every other: parsing
            // `:nth-child(2)` and then ignoring it would match every child.
            i += 1;
            if i < chars.len() && chars[i] == ':' {
                // `::before` — a pseudo-ELEMENT is a generated box, not a
                // condition on this one. Still refused.
                return None;
            }
            let start = i;
            while i < chars.len() && is_ident_char(chars[i]) {
                i += 1;
            }
            if start == i {
                return None;
            }
            let name: String = chars[start..i].iter().collect::<String>().to_ascii_lowercase();
            // A functional pseudo-class carries its argument in parentheses.
            let argument = if i < chars.len() && chars[i] == '(' {
                let open = i;
                let close = chars[open..].iter().position(|c| *c == ')')? + open;
                i = close + 1;
                Some(chars[open + 1..close].iter().collect::<String>())
            } else {
                None
            };
            let fixed = |from_end: bool, of_type: bool| SimplePart::NthChild {
                step: 0,
                offset: 1,
                from_end,
                of_type,
            };
            match (name.as_str(), argument) {
                // The shorthands ARE `an+b`, so they become it here and there
                // is only ever one matcher to be right. `-of-type` differs from
                // `-child` in WHAT is counted, never in how.
                ("first-child", None) => parts.push(fixed(false, false)),
                ("last-child", None) => parts.push(fixed(true, false)),
                ("first-of-type", None) => parts.push(fixed(false, true)),
                ("last-of-type", None) => parts.push(fixed(true, true)),
                // `:only-*` is first AND last — two conditions on one element,
                // which is exactly what a compound selector already is.
                ("only-child", None) => {
                    parts.push(fixed(false, false));
                    parts.push(fixed(true, false));
                }
                ("only-of-type", None) => {
                    parts.push(fixed(false, true));
                    parts.push(fixed(true, true));
                }
                ("nth-child", Some(arg))
                | ("nth-last-child", Some(arg))
                | ("nth-of-type", Some(arg))
                | ("nth-last-of-type", Some(arg)) => {
                    let (step, offset) = parse_nth(&arg)?;
                    parts.push(SimplePart::NthChild {
                        step,
                        offset,
                        from_end: name.starts_with("nth-last"),
                        of_type: name.ends_with("of-type"),
                    });
                }
                // `:not(…)` takes a selector LIST, and each alternative is a
                // compound — no combinators, which Selectors 3 also forbids and
                // this parser enforces by parsing compounds rather than
                // complexes.
                ("not", Some(arg)) => {
                    let mut inner = Vec::new();
                    for alternative in arg.split(',') {
                        let chars: Vec<char> = alternative.trim().chars().collect();
                        if chars.is_empty() {
                            return None;
                        }
                        let (compound, used) = parse_compound(&chars)?;
                        // Anything left over is a combinator or junk; either
                        // way this is not a selector we agreed to support.
                        if used != chars.len() {
                            return None;
                        }
                        inner.push(compound);
                    }
                    parts.push(SimplePart::Not(inner));
                }
                (state, None) if is_supported_state(state) => {
                    parts.push(SimplePart::State(name));
                }
                _ => return None,
            }
        } else {
            // Unknown character — refused rather than skipped: a selector that
            // drops a condition matches MORE than it was asked to.
            return None;
        }
    }
    if parts.is_empty() {
        None
    } else {
        Some((CompoundSelector { parts }, i))
    }
}

fn parse_attr(chars: &[char]) -> Option<(SimplePart, usize)> {
    let mut i = 0;
    let start = i;
    while i < chars.len() && is_ident_char(chars[i]) {
        i += 1;
    }
    if start == i {
        return None;
    }
    let name: String = chars[start..i].iter().collect();
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    if i >= chars.len() {
        return None;
    }
    if chars[i] == ']' {
        return Some((
            SimplePart::Attr {
                name,
                op: AttrOp::Has,
                value: None,
            },
            i + 1,
        ));
    }
    let op = match chars[i] {
        '=' => {
            i += 1;
            AttrOp::Exact
        }
        '*' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Substring
        }
        '^' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Prefix
        }
        '$' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Suffix
        }
        '~' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Word
        }
        '|' if i + 1 < chars.len() && chars[i + 1] == '=' => {
            i += 2;
            AttrOp::Lang
        }
        _ => return None,
    };
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    if i >= chars.len() {
        return None;
    }
    let value = if chars[i] == '"' || chars[i] == '\'' {
        let quote = chars[i];
        i += 1;
        let start = i;
        while i < chars.len() && chars[i] != quote {
            i += 1;
        }
        if i >= chars.len() {
            return None;
        }
        let v: String = chars[start..i].iter().collect();
        i += 1;
        v
    } else {
        let start = i;
        while i < chars.len() && is_ident_char(chars[i]) {
            i += 1;
        }
        chars[start..i].iter().collect()
    };
    while i < chars.len() && chars[i].is_whitespace() {
        i += 1;
    }
    if i >= chars.len() || chars[i] != ']' {
        return None;
    }
    Some((
        SimplePart::Attr {
            name,
            op,
            value: Some(value),
        },
        i + 1,
    ))
}

fn is_ident_start(c: char) -> bool {
    c.is_ascii_alphabetic() || c == '_' || c == '-' || c == '\\' || (c as u32) > 0x7F
}

fn is_ident_char(c: char) -> bool {
    c.is_ascii_alphanumeric() || c == '_' || c == '-' || c == '\\' || (c as u32) > 0x7F
}

/// Does `element` match this selector?
///
/// **Right to left**, which is not an optimisation but the only tractable
/// direction: the subject of a selector is its LAST compound, so matching
/// starts from the element in hand and walks outwards, rather than searching
/// the tree for candidates.
pub fn matches<E: Element>(element: &E, selector: &ComplexSelector) -> bool {
    let Some((_, subject)) = selector.parts.last() else {
        return false;
    };
    if !matches_compound(element, subject) {
        return false;
    }
    match_ancestors(element, &selector.parts[..selector.parts.len() - 1], selector)
}

/// Walk the remaining compounds outwards from `element`.
fn match_ancestors<E: Element>(
    element: &E,
    remaining: &[(Combinator, CompoundSelector)],
    selector: &ComplexSelector,
) -> bool {
    let Some((_, compound)) = remaining.last() else {
        return true;
    };
    // The combinator BETWEEN this compound and the one to its right is stored
    // on the right-hand entry — so the relationship to check is the one
    // recorded on the compound we just satisfied.
    let combinator = selector.parts[remaining.len()].0;
    let rest = &remaining[..remaining.len() - 1];
    match combinator {
        // Any ancestor may satisfy it, so every failure has to keep walking —
        // `a b c` where the first `b` found is the wrong one is the case a
        // single step gets wrong.
        Combinator::Descendant => {
            let mut current = element.parent();
            while let Some(node) = current {
                if matches_compound(&node, compound) && match_ancestors(&node, rest, selector) {
                    return true;
                }
                current = node.parent();
            }
            false
        }
        Combinator::Child => match element.parent() {
            Some(parent) => {
                matches_compound(&parent, compound) && match_ancestors(&parent, rest, selector)
            }
            None => false,
        },
        Combinator::AdjacentSibling => match element.previous_sibling() {
            Some(prev) => {
                matches_compound(&prev, compound) && match_ancestors(&prev, rest, selector)
            }
            None => false,
        },
        Combinator::GeneralSibling => {
            let mut current = element.previous_sibling();
            while let Some(node) = current {
                if matches_compound(&node, compound) && match_ancestors(&node, rest, selector) {
                    return true;
                }
                current = node.previous_sibling();
            }
            false
        }
    }
}

fn matches_compound<E: Element>(element: &E, compound: &CompoundSelector) -> bool {
    compound.parts.iter().all(|part| match part {
        SimplePart::Universal => true,
        // Only the tree that owns the widgets can answer this one.
        SimplePart::State(name) => element.state(name),
        // Count the siblings on ONE side — the position is 1-based, so the
        // element itself is the `1`.
        SimplePart::NthChild {
            step,
            offset,
            from_end,
            of_type,
        } => {
            // `-of-type` counts only siblings sharing this element's tag, so
            // the tag is fetched once and compared per step rather than per
            // element pair.
            let tag = of_type.then(|| element.tag());
            let mut position = 1i32;
            let mut cursor = if *from_end {
                element.next_sibling()
            } else {
                element.previous_sibling()
            };
            while let Some(node) = cursor {
                let counts = match &tag {
                    Some(want) => node.tag().eq_ignore_ascii_case(want),
                    None => true,
                };
                if counts {
                    position += 1;
                }
                cursor = if *from_end {
                    node.next_sibling()
                } else {
                    node.previous_sibling()
                };
            }
            nth_matches(position, *step, *offset)
        }
        // Negation: none of the alternatives may match. Recurses through the
        // same matcher, so anything a compound can express is negatable.
        SimplePart::Not(alternatives) => !alternatives
            .iter()
            .any(|compound| matches_compound(element, compound)),
        // Tag names are folded on the way into the DOM, so the comparison is
        // made in the same case rather than assuming the author's.
        SimplePart::Type(name) => element.tag().eq_ignore_ascii_case(name),
        SimplePart::Id(id) => element.attribute("id").as_deref() == Some(id.as_str()),
        // `class` is a LIST, and a substring test would let `.btn` match
        // `class="btn-primary"`.
        SimplePart::Class(class) => element
            .attribute("class")
            .map(|value| value.split_whitespace().any(|c| c == class))
            .unwrap_or(false),
        SimplePart::Attr { name, op, value } => {
            let Some(actual) = element.attribute(name) else {
                return false;
            };
            let Some(expected) = value else {
                return *op == AttrOp::Has;
            };
            match op {
                AttrOp::Has => true,
                AttrOp::Exact => actual == *expected,
                AttrOp::Substring => !expected.is_empty() && actual.contains(expected.as_str()),
                AttrOp::Prefix => !expected.is_empty() && actual.starts_with(expected.as_str()),
                AttrOp::Suffix => !expected.is_empty() && actual.ends_with(expected.as_str()),
                AttrOp::Word => actual.split_whitespace().any(|w| w == expected),
                AttrOp::Lang => {
                    actual == *expected || actual.starts_with(&format!("{expected}-"))
                }
            }
        }
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A tree in three lines, so the matcher is tested rather than a DOM.
    struct Node {
        tag: String,
        attrs: Vec<(String, String)>,
        parent: Option<Box<Node>>,
        previous: Option<Box<Node>>,
    }

    impl Node {
        fn new(tag: &str) -> Node {
            Node {
                tag: tag.to_string(),
                attrs: Vec::new(),
                parent: None,
                previous: None,
            }
        }
        fn attr(mut self, name: &str, value: &str) -> Node {
            self.attrs.push((name.to_string(), value.to_string()));
            self
        }
        fn under(mut self, parent: Node) -> Node {
            self.parent = Some(Box::new(parent));
            self
        }
        fn after(mut self, previous: Node) -> Node {
            self.previous = Some(Box::new(previous));
            self
        }
    }

    impl Element for &Node {
        fn tag(&self) -> String {
            self.tag.clone()
        }
        fn attribute(&self, name: &str) -> Option<String> {
            self.attrs
                .iter()
                .find(|(n, _)| n == name)
                .map(|(_, v)| v.clone())
        }
        fn parent(&self) -> Option<Self> {
            self.parent.as_deref()
        }
        fn previous_sibling(&self) -> Option<Self> {
            self.previous.as_deref()
        }
    }

    fn sel(s: &str) -> ComplexSelector {
        parse_selector_list(s).expect("parses").remove(0)
    }

    fn hits(node: &Node, s: &str) -> bool {
        matches(&node, &sel(s))
    }

    /// **The positional family and `:not()` MATCH**, not merely parse.
    ///
    /// Parsing proves only that the syntax was accepted; these are the
    /// assertions that would catch a matcher counting the wrong side or a
    /// negation that inverted nothing.
    #[test]
    fn positional_and_negation_match_what_they_name() {
        // `p` is third among its siblings: <span>, <p>, <p:this>.
        let first = Node::new("span");
        let second = Node::new("p").after(first);
        let third = Node::new("p").attr("class", "lead").after(second);

        assert!(hits(&third, "p:nth-child(3)"));
        assert!(!hits(&third, "p:nth-child(2)"));
        assert!(hits(&third, "p:nth-child(odd)"));
        // …but only the SECOND of its type, because a `<span>` came first and
        // `-of-type` does not count it.
        assert!(hits(&third, "p:nth-of-type(2)"));
        assert!(!hits(&third, "p:first-of-type"));
        // The `<span>` is first on both counts.
        let alone = Node::new("span");
        assert!(hits(&alone, "span:first-child"));
        assert!(hits(&alone, "span:first-of-type"));

        // Negation, including over a pseudo-class — the recursion is what
        // makes that free.
        assert!(hits(&third, "p:not(.intro)"));
        assert!(!hits(&third, "p:not(.lead)"));
        assert!(!hits(&alone, "span:not(:first-child)"));
        // A list means NEITHER.
        assert!(!hits(&third, "p:not(.lead, span)"));
        assert!(hits(&third, "p:not(.intro, span)"));
    }

    #[test]
    fn the_atoms_match_what_they_name() {
        let node = Node::new("a").attr("id", "home").attr("class", "btn wide");
        assert!(hits(&node, "a"));
        assert!(hits(&node, "*"));
        assert!(hits(&node, "#home"));
        assert!(hits(&node, ".btn"));
        assert!(hits(&node, ".wide"));
        assert!(hits(&node, "a#home.btn.wide"));
        assert!(!hits(&node, "div"));
        assert!(!hits(&node, "#away"));
    }

    #[test]
    fn a_class_is_a_list_not_a_substring() {
        // `.btn` matching `class="btn-primary"` is the classic wrong answer,
        // and a `contains` is how it happens.
        let node = Node::new("a").attr("class", "btn-primary");
        assert!(!hits(&node, ".btn"));
        assert!(hits(&node, ".btn-primary"));
    }

    #[test]
    fn every_attribute_operator_means_something_different() {
        let node = Node::new("input")
            .attr("type", "text")
            .attr("class", "a b")
            .attr("lang", "en-GB");
        assert!(hits(&node, "[type]"));
        assert!(hits(&node, "[type=text]"));
        assert!(hits(&node, "[type*=ex]"));
        assert!(hits(&node, "[type^=te]"));
        assert!(hits(&node, "[type$=xt]"));
        assert!(hits(&node, "[class~=b]"), "one of a whitespace list");
        // Quoted, because an unquoted value cannot contain a space — and it
        // can never match, since `~=` compares against single words.
        assert!(!hits(&node, "[class~=\"a b\"]"));
        assert!(hits(&node, "[class=\"a b\"]"), "but `=` is the whole value");
        assert!(hits(&node, "[lang|=en]"), "en-GB is en, hyphen-prefixed");
        assert!(!hits(&node, "[lang|=e]"), "and not any old prefix");
        assert!(!hits(&node, "[disabled]"));
    }

    #[test]
    fn combinators_ask_about_different_relatives() {
        let section = Node::new("section").attr("id", "main");
        let div = Node::new("div").under(section);
        let p = Node::new("p").under(div);

        assert!(hits(&p, "#main p"), "descendant crosses any depth");
        assert!(!hits(&p, "#main > p"), "child does not");
        assert!(hits(&p, "div > p"));

        let h = Node::new("h1");
        let after = Node::new("p").after(h);
        assert!(hits(&after, "h1 + p"));
        assert!(hits(&after, "h1 ~ p"));
        let far = Node::new("p").after(Node::new("span").after(Node::new("h1")));
        assert!(!hits(&far, "h1 + p"), "adjacent means IMMEDIATELY before");
        assert!(hits(&far, "h1 ~ p"), "general sibling does not");
    }

    #[test]
    fn a_descendant_step_keeps_looking_after_a_near_miss() {
        // `body div p` where the nearest `div` is inside another `div` that is
        // NOT under a body — a single step outwards gives up too early.
        let body = Node::new("body");
        let outer = Node::new("div").under(body);
        let inner = Node::new("div").under(outer);
        let p = Node::new("p").under(inner);
        assert!(hits(&p, "body div p"));
    }

    #[test]
    fn specificity_orders_the_way_the_cascade_needs() {
        assert!(sel("#a").specificity() > sel(".a").specificity());
        assert!(sel(".a").specificity() > sel("a").specificity());
        assert!(sel("[href]").specificity() == sel(".a").specificity());
        // Counted across the whole chain, so a descendant selector outranks the
        // compound it ends with.
        assert!(sel("div p").specificity() > sel("p").specificity());
        // The universal selector contributes nothing.
        assert_eq!(sel("*").specificity(), 0);
    }

    #[test]
    fn an_unsupported_selector_is_refused_rather_than_narrowed() {
        // A pseudo-class that parsed and then never matched would be a rule
        // that silently applies to everything the rest of it allows. So the
        // line is drawn at what this engine can ANSWER, not at what it can
        // tokenise.
        assert!(parse_selector_list("a:hover").is_some());
        assert!(parse_selector_list("input:checked").is_some());
        // …and everything it cannot answer is still refused whole. `:focus`
        // and `:disabled` look every bit as available as `:hover` and are not:
        // the form tracks focus by child index, and `SetEnabled` has no `Get`.
        assert!(parse_selector_list("a:focus").is_none());
        assert!(parse_selector_list("button:disabled").is_none());
        assert!(parse_selector_list("p::before").is_none());
        // …and the structural ones are answerable now: they are counted from
        // the sibling chain, which the tree already exposes.
        assert!(parse_selector_list("li:nth-child(2)").is_some());
        assert!(parse_selector_list("li:first-child").is_some());
        assert!(parse_selector_list("li:nth-child(odd)").is_some());
        // A malformed argument is still a refusal, not a rule that matches
        // nothing quietly.
        assert!(parse_selector_list("li:nth-child(banana)").is_none());
        // `:not()` and the `-of-type` family.
        assert!(parse_selector_list("p:not(.lead)").is_some());
        assert!(parse_selector_list("p:not(.a, #b)").is_some());
        assert!(parse_selector_list("p:first-of-type").is_some());
        assert!(parse_selector_list("p:nth-of-type(2n)").is_some());
        // A combinator inside `:not()` is not a compound — Selectors 3 forbids
        // it and so does this parser, rather than matching something narrower
        // than the author wrote.
        assert!(parse_selector_list("p:not(div p)").is_none());
        assert!(parse_selector_list("p:not()").is_none());
        assert!(parse_selector_list("li:nth-child(").is_none());
        // A bare colon, or one with nothing after it, is a syntax error rather
        // than a state.
        assert!(parse_selector_list("a:").is_none());
        assert!(parse_selector_list("").is_none());
        assert!(parse_selector_list("a,").is_none());
    }
}
