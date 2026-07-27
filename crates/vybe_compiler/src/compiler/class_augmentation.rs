//! Class augmentation — where a class's members come from besides its own body.
//!
//! PHP `trait`/`use`, Dart `mixin`/`with`, Ruby `include`/`prepend`, Java
//! interface default methods, Go field promotion, Dart `extension E on MyClass`.
//! Five languages implement this today, five separate times, all in walkers.
//! This is the one model that replaces them. See flexclassplan.md §4c.
//!
//! ## One vocabulary, per-language data — NOT one algorithm
//!
//! These mechanisms differ in KIND, so a single fold would be wrong for most:
//!
//! - **Copy** — members are duplicated into the class (PHP traits, Dart mixins).
//! - **Chain** — members are inserted into the LOOKUP ORDER, not copied
//!   (Ruby `include`/`prepend`; Java defaults resolve at dispatch).
//! - **Promote** — members are promoted from an inner value and THE RECEIVER
//!   REBINDS to it (Go field promotion).
//!
//! A language declares `Augmentation` records on its `NormalClass`; this pass
//! applies them once. A language that declares none is untouched, so migration
//! is one language at a time with no flag day.
//!
//! ## Ordering
//!
//! This runs on the normalized class, BEFORE member registration. Contributed
//! members must exist by the time `predeclare_class_surface` (link.rs) records
//! the class's member list — otherwise they are invisible to receiver-typed
//! resolution and the compilation-order bug returns by another route (§3a).
//!
//! ## Errors are loud
//!
//! Go promotion at equal depth and Java default-method diamonds are errors in
//! the source language. A last-one-wins fold hides them; `Conflict::Error` and
//! `RequireExplicit` surface them.

use std::collections::HashMap;

use vybe_bytecode::class_normalize::{
    Augmentation, AugmentationConflict, AugmentationMode, AugmentationPosition, NormalClass,
    NormalMethod,
};

/// A conflict the augmentation pass could not resolve on its own.
#[derive(Debug, Clone)]
pub struct AugmentationError {
    pub class: String,
    pub member: String,
    pub sources: Vec<String>,
    pub reason: &'static str,
}

impl std::fmt::Display for AugmentationError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(
            f,
            "class `{}`: member `{}` {} (from {})",
            self.class,
            self.member,
            self.reason,
            self.sources.join(", ")
        )
    }
}

/// Resolve one class's declared augmentations into its member list.
///
/// `available` supplies the augmenting types by name — the same normalized
/// classes, so a trait that uses a trait or a mixin on a mixin resolves
/// transitively as long as its own augmentations were applied first.
///
/// Returns the conflicts that the declared policy says are errors. The caller
/// decides how to report them; this pass never picks a winner behind the
/// user's back.
pub fn apply_augmentations(
    class: &mut NormalClass,
    available: &HashMap<String, NormalClass>,
) -> Vec<AugmentationError> {
    if class.augmentations.is_empty() {
        return Vec::new();
    }

    let mut errors = Vec::new();
    // Members the class declares itself. These are never overwritten by an
    // `AfterOwn` augmentation — every language agrees the class body wins.
    let own: Vec<String> = class
        .instance_methods
        .iter()
        .map(|m| m.canonical_name.clone())
        .collect();

    // Which augmentation supplied each contributed name, for conflict reporting.
    let mut supplied_by: HashMap<String, Vec<String>> = HashMap::new();
    let mut depth_of: HashMap<String, u8> = HashMap::new();

    let augmentations = class.augmentations.clone();
    for aug in &augmentations {
        let Some(source) = available.get(&aug.from) else {
            // An augmenting type the compiler cannot see (declared in another
            // module, or a framework type). Not an error here — the class's
            // member list is simply partial, which callers already handle.
            continue;
        };

        for method in &source.instance_methods {
            let Some(name) = adjusted_name(aug, &method.canonical_name) else {
                continue; // excluded by `insteadof`
            };

            if !aug.contributes.methods {
                continue;
            }

            // The class's own declaration wins unless the augmentation is
            // positioned before it (Ruby `prepend`).
            if own.contains(&name) && aug.position == AugmentationPosition::AfterOwn {
                continue;
            }

            // `Chain` is NOT a copy (§4c). Ruby `prepend` inserts the module
            // AHEAD of the class in the lookup order, and `super` inside the
            // prepended method reaches the class's own — shadowed, not
            // replaced. Copying over the class's member would delete the very
            // thing `super` must find, so refuse rather than silently
            // mis-compile. Ruby exercises `super` through a module in 23 files;
            // this would be caught immediately, but a loud refusal is the
            // contract (§2f).
            if aug.mode == AugmentationMode::Chain && own.contains(&name) {
                errors.push(AugmentationError {
                    class: class.name.clone(),
                    member: name.clone(),
                    sources: vec![aug.from.clone()],
                    reason: "chain-order insertion over an existing member is not implemented \
                             (a copy would delete the member `super` must reach)",
                });
                continue;
            }

            let previous = supplied_by.entry(name.clone()).or_default();
            if !previous.is_empty() {
                match aug.conflict {
                    AugmentationConflict::FirstWins => {
                        previous.push(aug.from.clone());
                        continue;
                    }
                    AugmentationConflict::Error => {
                        previous.push(aug.from.clone());
                        errors.push(AugmentationError {
                            class: class.name.clone(),
                            member: name.clone(),
                            sources: previous.clone(),
                            reason: "supplied by more than one augmentation",
                        });
                        continue;
                    }
                    AugmentationConflict::RequireExplicit => {
                        previous.push(aug.from.clone());
                        errors.push(AugmentationError {
                            class: class.name.clone(),
                            member: name.clone(),
                            sources: previous.clone(),
                            reason: "ambiguous; the class must resolve it explicitly",
                        });
                        continue;
                    }
                    AugmentationConflict::LastWins => {}
                }
            }

            // Go promotion: shallower depth wins outright; EQUAL depth with the
            // same name is ambiguous, which is an error in Go rather than a
            // silent pick.
            if aug.mode == AugmentationMode::Promote {
                match depth_of.get(&name).copied() {
                    Some(seen) if seen < aug.depth => continue,
                    Some(seen) if seen == aug.depth => {
                        previous.push(aug.from.clone());
                        errors.push(AugmentationError {
                            class: class.name.clone(),
                            member: name.clone(),
                            sources: previous.clone(),
                            reason: "promoted at equal depth from more than one field",
                        });
                        continue;
                    }
                    _ => {}
                }
                depth_of.insert(name.clone(), aug.depth);
            }

            previous.push(aug.from.clone());
            class
                .instance_methods
                .retain(|existing| existing.canonical_name != name);
            class.instance_methods.push(contributed(aug, method, &name));
        }

        // Fields. A class's own field of the same name always wins; there is
        // no language in which an augmenting type's field shadows a declared
        // one.
        if aug.contributes.fields {
            for field in &source.instance_fields {
                if class.instance_fields.iter().any(|f| f.name == field.name) {
                    continue;
                }
                class.instance_fields.push(field.clone());
            }
        }

        // Properties (getter/setter pairs) travel with methods — a Dart mixin
        // or PHP trait contributing `int get area => …` must bring it, or the
        // member is silently missing from the augmented class.
        if aug.contributes.methods {
            for property in &source.properties {
                if class
                    .properties
                    .iter()
                    .any(|p| p.canonical_name == property.canonical_name)
                {
                    continue;
                }
                class.properties.push(property.clone());
            }
        }

        // Statics. PHP quirk: a trait's static property gives each using class
        // its OWN copy, so this is a copy, never a shared reference.
        if aug.contributes.statics {
            for field in &source.static_fields {
                if class.static_fields.iter().any(|f| f.name == field.name) {
                    continue;
                }
                class.static_fields.push(field.clone());
            }
            for method in &source.static_methods {
                if class
                    .static_methods
                    .iter()
                    .any(|m| m.canonical_name == method.canonical_name)
                {
                    continue;
                }
                class.static_methods.push(method.clone());
            }
        }

        // `Chain` and `Promote` also record the source for identity checks
        // (`is` / `instanceof` / `kind_of?`), which walk `interfaces`.
        if !class.interfaces.iter().any(|i| i == &aug.from) {
            class.interfaces.push(aug.from.clone());
        }
    }

    errors
}

/// The name a member is bound under after this augmentation's adjustments,
/// or `None` when it is excluded (PHP `insteadof`).
fn adjusted_name(aug: &Augmentation, member: &str) -> Option<String> {
    for adj in &aug.adjustments {
        if adj.member != member {
            continue;
        }
        if adj.exclude {
            return None;
        }
        if let Some(renamed) = &adj.rename_to {
            return Some(renamed.clone());
        }
    }
    Some(member.to_string())
}

/// Build the contributed method, applying the augmentation's name and
/// visibility adjustments.
fn contributed(aug: &Augmentation, method: &NormalMethod, name: &str) -> NormalMethod {
    let mut out = method.clone();
    out.canonical_name = name.to_string();
    for adj in &aug.adjustments {
        if adj.member == method.canonical_name {
            if let Some(visibility) = adj.visibility {
                out.access = visibility;
            }
        }
    }
    out
}
