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
//! ## Deliberately absent
//!
//! **Pseudo-classes and pseudo-elements.** `:hover`, `:nth-child()`,
//! `::before`. They need either live interaction state or generated boxes, and
//! a selector that parses and then silently never matches is worse than one
//! that is refused — so an unknown character makes the whole selector `None`
//! rather than a selector that quietly drops a condition.

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
    pub fn specificity(&self) -> u32 {
        let (mut ids, mut classes, mut types) = (0u32, 0u32, 0u32);
        for (_, compound) in &self.parts {
            for part in &compound.parts {
                match part {
                    SimplePart::Id(_) => ids += 1,
                    SimplePart::Class(_) | SimplePart::Attr { .. } => classes += 1,
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
}

/// Parse a comma-separated selector list. `None` if any of it is unsupported.
pub fn parse_selector_list(input: &str) -> Option<Vec<ComplexSelector>> {
    let mut out = Vec::new();
    for piece in input.split(',') {
        let trimmed = piece.trim();
        if trimmed.is_empty() {
            return None;
        }
        out.push(parse_complex(trimmed)?);
    }
    if out.is_empty() { None } else { Some(out) }
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
        } else {
            // Unknown character — a pseudo-class, most likely. Refused rather
            // than skipped: a selector that drops a condition matches MORE than
            // it was asked to.
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
        // that silently applies to everything the rest of it allows.
        assert!(parse_selector_list("a:hover").is_none());
        assert!(parse_selector_list("p::before").is_none());
        assert!(parse_selector_list("li:nth-child(2)").is_none());
        assert!(parse_selector_list("").is_none());
        assert!(parse_selector_list("a,").is_none());
    }
}
