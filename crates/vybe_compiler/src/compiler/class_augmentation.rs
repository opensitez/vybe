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
    Augmentation, AugmentationAdjustment, AugmentationConflict, AugmentationMode,
    AugmentationPosition, NormalClass, NormalMethod,
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
    // Members the class IMPLEMENTS itself. These are never overwritten by an
    // `AfterOwn` augmentation — every language agrees the class body wins.
    //
    // An ABSTRACT declaration is not an implementation, it is a REQUIREMENT
    // ("someone must supply this"), and a contributed concrete member satisfies
    // it rather than losing to it. Counting it as own leaves the class with a
    // bodiless method and silently drops the implementation. Every language
    // with both features agrees: a PHP trait method satisfies an `abstract
    // function`, a Java interface `default` satisfies an abstract method, a
    // Dart mixin member satisfies an abstract one.
    // Instance and static are separate surfaces: a class may declare a static
    // `make` and receive an instance `make`, and neither shadows the other.
    let own: Vec<(bool, String)> = class
        .instance_methods
        .iter()
        .map(|m| (false, m))
        .chain(class.static_methods.iter().map(|m| (true, m)))
        .filter(|(_, m)| !m.is_abstract)
        .map(|(is_static, m)| (is_static, m.canonical_name.clone()))
        .collect();

    // Which augmentation supplied each contributed name, for conflict reporting.
    let mut supplied_by: HashMap<(bool, String), Vec<String>> = HashMap::new();
    let mut depth_of: HashMap<String, u8> = HashMap::new();

    // The class's own name, taken before the member loops borrow `class`
    // mutably — a promoted forwarder needs it for the receiver's type.
    let class_name = class.name.clone();
    let augmentations = class.augmentations.clone();
    for aug in &augmentations {
        let Some(source) = available.get(&aug.from) else {
            // An augmenting type the compiler cannot see (declared in another
            // module, or a framework type). Not an error here — the class's
            // member list is simply partial, which callers already handle.
            continue;
        };

        // Both method surfaces, under the SAME rules. A rename, an exclusion or
        // a visibility change applies to whatever the trait declared — PHP's
        // `use A { helper as protected make; }` is legal whether `helper` is
        // static or not, and handling only instance methods silently ignored
        // every adaptation on a static one.
        let sources: [(bool, &Vec<NormalMethod>); 2] = [
            (false, &source.instance_methods),
            (true, &source.static_methods),
        ];
        for (is_static, source_methods) in sources {
            let permitted = if is_static {
                aug.contributes.statics
            } else {
                aug.contributes.methods
            };
            if !permitted {
                continue;
            }
            for method in source_methods {
                // A source member may bind under SEVERAL names. PHP's `as` is
                // additive — `use A { hello as hi; }` gives the class BOTH `hello`
                // and `hi` — and it composes with `insteadof`, which is the whole
                // point of `B::hello insteadof A; A::hello as helloFromA;`: hide the
                // conflicting name, keep the implementation reachable under another.
                // A rename that REPLACED the original could express neither.
                for (name, adjustment) in bound_names(aug, &method.canonical_name) {
                    let bound = (is_static, name.clone());
                    // The class's own declaration wins unless the augmentation is
                    // positioned before it (Ruby `prepend`).
                    if own.contains(&bound) && aug.position == AugmentationPosition::AfterOwn {
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
                    if aug.mode == AugmentationMode::Chain && own.contains(&bound) {
                        errors.push(AugmentationError {
                        class: class.name.clone(),
                        member: name.clone(),
                        sources: vec![aug.from.clone()],
                        reason: "chain-order insertion over an existing member is not implemented \
                             (a copy would delete the member `super` must reach)",
                    });
                        continue;
                    }

                    let previous = supplied_by.entry(bound.clone()).or_default();
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
                    // The member's ROLE travels with it. A trait supplying
                    // `__toString` gives the using class a ToString, and a
                    // mixin supplying `operator +` gives it a Plus — dropping
                    // the role here would contribute the body but leave the
                    // class with no slot, so a cross-language call would miss a
                    // method the class demonstrably has. Only under the
                    // member's OWN name: an alias is a second entry point, not
                    // a second implementation of the role.
                    if name == method.canonical_name {
                        if let Some(role) = source
                            .special_methods
                            .iter()
                            .find(|s| s.canonical_name == method.canonical_name)
                        {
                            if !class
                                .special_methods
                                .iter()
                                .any(|s| s.canonical_name == name)
                            {
                                class.special_methods.push(role.clone());
                            }
                        }
                    }
                    let member = match (&aug.mode, aug.via_field.as_deref()) {
                        // `Promote` is NOT a copy. Go's promoted method runs on
                        // the INNER value — `outer.M()` is `outer.f.M()` — so
                        // copying the body would leave it operating on the
                        // outer receiver, which is a different struct with
                        // different fields. What the class gains is a
                        // FORWARDER; the rebinding is the whole point of the
                        // mode.
                        (AugmentationMode::Promote, Some(field)) => {
                            promoted(method, &name, field, &class_name, adjustment.as_ref())
                        }
                        _ => contributed(method, &name, adjustment.as_ref()),
                    };
                    let target = if is_static {
                        &mut class.static_methods
                    } else {
                        &mut class.instance_methods
                    };
                    target.retain(|existing| existing.canonical_name != name);
                    target.push(member);
                }
            }
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

        // Static FIELDS. PHP quirk: a trait's static property gives each using
        // class its OWN copy, so this is a copy, never a shared reference.
        // Static METHODS are contributed by the method loop above, under the
        // same adjustments as instance ones.
        if aug.contributes.statics {
            for field in &source.static_fields {
                if class.static_fields.iter().any(|f| f.name == field.name) {
                    continue;
                }
                class.static_fields.push(field.clone());
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

/// Every name this member binds under through this augmentation, each paired
/// with the adjustment that produced it (`None` for the member's own name).
///
/// The default name is present unless an `exclude` adjustment covers it (PHP
/// `insteadof`), PLUS one entry per `rename_to` (PHP `as`). Both may apply at
/// once — an excluded member stays reachable under its alias, which is the
/// documented PHP behaviour and what four of the trait tests assert.
///
/// An augmentation with no adjustments yields exactly the member's own name, so
/// a language that declares none (Dart today) is unaffected.
fn bound_names(aug: &Augmentation, member: &str) -> Vec<(String, Option<AugmentationAdjustment>)> {
    let mut names = Vec::new();
    let mut excluded = false;
    for adj in &aug.adjustments {
        if adj.member != member {
            continue;
        }
        if adj.exclude {
            excluded = true;
        }
        if let Some(renamed) = &adj.rename_to {
            names.push((renamed.clone(), Some(adj.clone())));
        }
    }
    if !excluded {
        // The visibility-only form (`A::run as protected;`) adjusts the member
        // under its own name, so it has to travel with the default binding.
        let own = aug
            .adjustments
            .iter()
            .find(|adj| adj.member == member && adj.rename_to.is_none() && !adj.exclude)
            .cloned();
        names.insert(0, (member.to_string(), own));
    }
    names
}

/// Build the contributed method under one bound name.
///
/// Visibility comes from the adjustment that produced THIS binding, never from
/// a sibling one: `A::run as protected runP;` makes the alias protected and
/// leaves `run` itself alone.
/// A PROMOTED member: a forwarder onto the inner value, not a copy of the body.
///
/// `func (o Outer) M(a) R { return o.<field>.M(a) }`
///
/// The receiver rebinds — which is what separates `Promote` from `Copy`. The
/// inner call names the source member's OWN spelling, because that is the name
/// it is stored under on the inner value; only the outer entry point takes the
/// (possibly renamed) bound name.
fn promoted(
    method: &NormalMethod,
    name: &str,
    via_field: &str,
    outer_class: &str,
    adjustment: Option<&AugmentationAdjustment>,
) -> NormalMethod {
    use vybe_ast::{Argument, ExprKind, Expression, Statement, StmtKind};

    let mut out = contributed(method, name, adjustment);
    // The receiver is a declared parameter in every language that promotes
    // (Go writes it `func (o Outer) M()`), so params[0] IS the receiver and
    // its type is now the OUTER struct.
    let receiver = out
        .params
        .first()
        .map(|param| param.name.clone())
        .unwrap_or_else(|| "self".to_string());
    if let Some(param) = out.params.first_mut() {
        param.type_hint = Some(outer_class.to_string());
    }
    let args: Vec<Argument> = out
        .params
        .iter()
        .skip(1)
        .map(|param| Argument::positional(Expression::ident(&param.name)))
        .collect();
    let inner = Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(&receiver)),
        field: via_field.to_string(),
        null_safe: false,
    });
    let call = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(inner),
            field: method.source_name.clone(),
            null_safe: false,
        })),
        args,
        optional: false,
    });
    out.body = vec![Statement::new(StmtKind::Return(Some(call)))];
    out
}

fn contributed(
    method: &NormalMethod,
    name: &str,
    adjustment: Option<&AugmentationAdjustment>,
) -> NormalMethod {
    let mut out = if name == method.canonical_name {
        method.clone()
    } else {
        // The model owns what "bound under another name" means; this pass only
        // says WHICH name.
        method.bound_as(name)
    };
    if let Some(visibility) = adjustment.and_then(|adj| adj.visibility) {
        out.access = visibility;
    }
    out
}
