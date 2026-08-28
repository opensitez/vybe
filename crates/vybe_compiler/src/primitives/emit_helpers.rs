//! Assorted lowering helpers: component-model, type predicates, Fortran ctors, global maps, records, VB statements.
//!
//! Extracted from `primitives/mod.rs` (`impl Compiler`) — conductor pattern,
//! same as `statements.rs`/`builtins.rs`.

use crate::primitives::class_slots;
use super::*;

impl Compiler {
    #[allow(dead_code)]
    pub(super) fn current_offset(&self) -> usize {
        self.chunks[self.current].current_offset()
    }
    pub(crate) fn str_const(&mut self, s: &str) -> u16 {
        self.chunks[self.current].add_constant(Value::String(Arc::from(s)))
    }

    pub(crate) fn import(&mut self, module: &str, name: &str) -> u16 {
        self.chunks[self.current].add_import(module, name)
    }

    /// Emit `ref.test` / `ref.cast` against the type spelled `name`.
    ///
    /// The name is resolved to a HEAPTYPE here, at compile time — abstract for
    /// the spec's own spellings, otherwise an index into the module's type
    /// section. The instruction never carries the name; only the type section
    /// does.
    pub(crate) fn emit_ref_type_test(&mut self, op: Op, name: &str, line: u32) {
        // ⚠ AN ARRAY TYPE LIVES UNDER A DIFFERENT ROW THAN ITS OWN NAME.
        //
        // `array.new`/`array.new_fixed $t` stamp the rtt from the row
        // `__wast_array::$t` (`resolve_gc_array_type_id`), while
        // `heaptype_for_name` below reserves a row named `$t` — two entries,
        // two ids, for one declared type. So `ref.test $t` on an array
        // allocated as `$t` compared a type against itself and answered FALSE:
        //
        //   (type $s (sub (array i8)))
        //   (ref.test (ref null $s) (array.new_fixed $s 0))   ;; was 0, must be 1
        //
        // Preferring the array row when one exists makes both spellings name
        // the same type. It cannot capture a struct: the key is only ever
        // created by the array registration path.
        let array_key = format!("__wast_array::{name}");
        if let Some(idx) = self.chunks[0].types.iter().position(|t| t.name == array_key) {
            let ht = vybe_runtime::opcode::heaptype::HeapType::Concrete(idx as u32 + 1);
            self.chunks[self.current].emit_ref_type_op(op, ht, line);
            return;
        }
        let ht = crate::primitives::classes::heaptype_for_name(&mut self.chunks, name);
        self.chunks[self.current].emit_ref_type_op(op, ht, line);
    }
    /// Emit a host call.
    ///
    /// Arity against a DECLARED signature is checked in `Chunk::emit_call`, not
    /// here: 181 sites emit a host call without going through this helper, so a
    /// check here covered barely half of them — including `createElement`,
    /// which `emit_control_element` emits directly.
    pub(crate) fn emit_host_call(&mut self, idx: u16, argc: u8) {
        let l = self.line;
        self.chunks[self.current].emit_call(idx, argc, l);
    }

    /// Resolve a qualified identifier to a Component Model host call
    /// `(module, function)` pair when its first segment matches the
    /// profile's `host_packages` list, else `None`.
    ///
    /// Walker conventions: PHP passes backslash-separated names
    /// (`Vybe\Http\Request\method`), other languages should normalize
    /// their separator to `\` before this point (TODO for Python / C# /
    /// etc.). This keeps the resolver language-agnostic.
    ///
    /// Mapping:
    /// - `[Vybe, Http, Request, method]` → `("vybe:http/request", "method")`
    /// - `[Vybe, Math, cos]`             → `("ecma:math", "cos")`
    /// - `[Wasi, Cli, log]`              → `("wasi:logging/logging", "log")`
    ///
    /// First join is `:` (package → interface), further joins use `/`,
    /// last segment is the function name. Everything is lowercased.
    pub(super) fn resolve_component_model_call(&self, name: &str) -> Option<(String, String)> {
        if !name.contains('\\') {
            return None;
        }
        let parts: Vec<&str> = name.split('\\').collect();
        if parts.len() < 2 {
            return None;
        }

        // namespaceplan.md: the global namespace tree is the PRIMARY
        // resolver — it mounts every host export (`vybe.gui.*` next to
        // `ecma.*` and `wasi.*`), so a backslash chain resolves exactly
        // like the dotted chains other languages emit. The manual
        // module-string build below remains the fallback for name shapes
        // the tree doesn't key (e.g. CamelCase segments of kebab-case
        // module names).
        if self.profile.uses_common_resolver {
            match self.resolve_namespace_path(&parts) {
                Some(self::resolver::Resolution::HostImport { module, func }) => {
                    return Some((module, func));
                }
                Some(self::resolver::Resolution::Tree(
                    crate::primitives::namespaces::ResolutionTarget::HostCall {
                        module, func, ..
                    },
                )) => {
                    return Some((module, func));
                }
                _ => {}
            }
        }

        // Consult the Linker's `host_package_roots` map instead of
        // `profile.namespaces.host_packages`. Populated at link time
        // from `EsmDefault::PackageRoot` entries (which the profile
        // loader auto-translates from the legacy list). Component
        // Model package names are lowercase by spec — match
        // case-insensitively regardless of the language's case rules.
        let first_key = parts[0].to_ascii_lowercase();
        if !self.host_package_roots.contains_key(&first_key) {
            return None;
        }

        let lower: Vec<String> = parts.iter().map(|s| s.to_ascii_lowercase()).collect();
        let (func, path) = lower.split_last()?;
        if path.is_empty() {
            return None;
        }

        let module = if path.len() == 1 {
            path[0].clone()
        } else {
            let mut m = path[0].clone();
            m.push(':');
            m.push_str(&path[1]);
            for p in &path[2..] {
                m.push('/');
                m.push_str(p);
            }
            m
        };
        Some((module, func.clone()))
    }

    // ── Crate-private accessors used by `dotnet_register` ──────────────
    //
    // The .NET class registration logic lives in a sibling file
    // (`dotnet_register.rs`) but operates on Compiler internals. These
    // helpers expose just the bits that registration needs without
    // making the underlying fields `pub`.
    #[allow(dead_code)]
    pub(crate) fn chunks_mut(&mut self) -> &mut Vec<Chunk> {
        &mut self.chunks
    }
    #[allow(dead_code)]
    pub(crate) fn current_line(&self) -> u32 {
        self.line
    }
    #[allow(dead_code)]
    pub(crate) fn note_defined_global(&mut self, name: &str) {
        self.defined_globals.insert(name.to_string());
    }
    #[allow(dead_code)]
    pub(crate) fn note_defined_class(&mut self, name: &str) {
        self.declare_class_identity(name);
    }
    /// Mount a namespace-tree root as ambient (unqualified names resolve under
    /// it) — used when a module imports a platform surface (`flutter.*`).
    pub(crate) fn mount_ambient_root(&mut self, root: &str) {
        if !self.ambient_tree_roots.iter().any(|r| r == root) {
            self.ambient_tree_roots.push(root.to_string());
        }
    }
    pub(crate) fn note_pending_class(&mut self, name: &str, parent: Option<String>) {
        self.pending_classes.insert(
            name.to_string(),
            PendingClass {
                bases: parent.iter().cloned().collect(),
                parent,
                enclosing_class: self.current_class.clone(),
                fields: Vec::new(),
                field_storage_names: HashMap::new(),
                is_value_type: false,
                instance_member_names: Vec::new(),
                instance_pointer_method_names: Vec::new(),
                instance_field_types: HashMap::new(),
                static_fields: Vec::new(),
                static_field_types: HashMap::new(),
                static_method_names: Vec::new(),
                instance_method_overloads: HashMap::new(),
                static_method_overloads: HashMap::new(),
                nested_types: Vec::new(),
                statics: Vec::new(),
            },
        );
    }

    /// Push the canonical event-registry key for a control expression.
    /// Used by AddHandler / RemoveHandler so the GUI host indexes handlers by
    /// the source-stable identifier (field name, class name for `Me`, etc.)
    /// rather than the runtime `.Name` property — renaming a control after
    /// the handler is wired must NOT break dispatch.
    ///
    pub(crate) fn canon(&self, name: &str) -> String {
        let name = self.variable_name_body(name);
        if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        }
    }

    /// `name` with any variable-namespace marker removed, so canonicalization
    /// folds the bare name. Identity for the languages that have a single
    /// namespace — which is all of them but PHP. See
    /// [`vybe_runtime::registry::VariableNamespace`].
    pub(crate) fn variable_name_body<'a>(&self, name: &'a str) -> &'a str {
        match self.variable_namespace {
            Some(ns) => (ns.body)(name),
            None => name,
        }
    }

    /// True when `name` is spelled in the language's separate VARIABLE
    /// namespace. Derived from `body` being non-identity so a language declares
    /// its marker in exactly one place instead of once per question asked
    /// about it.
    pub(crate) fn is_variable_name(&self, name: &str) -> bool {
        self.variable_name_body(name).len() != name.len()
    }

    /// Canonical global name for a *type* reference: [`Self::canon`] plus
    /// namespace-separator normalization.
    ///
    /// `\` (PHP) and `::` (C++ / Pascal / Ruby) both mean "namespace boundary";
    /// globals spell it `.`. Neither character can occur inside an identifier in
    /// any supported language, so normalizing both is separator-agnostic rather
    /// than a language check — which is what the scattered
    /// `canon(name).replace('\\', ".")` calls amount to today.
    pub(crate) fn canon_type_global(&self, name: &str) -> String {
        self.canon(name).replace("::", ".").replace('\\', ".")
    }

    /// Global-scope key for `name`, whose canonical body is `canon`.
    ///
    /// Where a language keeps variables in their own namespace, a global `$foo`
    /// must not collide with a function `foo`; the language decides what key
    /// separates them, and which names are exempt. Everywhere else the
    /// canonical name IS the key. Pairs with [`Self::canon`] — that strips the
    /// marker so this can key off the bare name, so the two change together.
    pub(crate) fn variable_global_key(&self, name: &str, canon: &str) -> String {
        self.variable_namespace
            .and_then(|ns| (ns.global_key)(name, canon))
            .unwrap_or_else(|| canon.to_string())
    }

    /// True if `name` is a class the program actually declares. A real class
    /// of a built-in exception name (e.g. PHP's prelude `LogicException`,
    /// `RuntimeException`, …) must go through the ordinary class emitter, NOT
    /// the `is_exception_type` intrinsic shortcut — otherwise the intrinsic
    /// shape (canonicalized `__type`, no `__types` chain) shadows the real
    /// class and subclass identity is lost.
    /// THE question "does a user declaration own this name", asked once.
    ///
    /// It was asked at least three different ways over the flat `defined_classes`
    /// set — a raw `contains`, a `contains(canon(..))`, and a case-insensitive
    /// linear scan — plus separate hand-rolled probes in the GUI paths. Three
    /// spellings of one question answer differently, which is how a user
    /// `Class Point` could win in the construction path and lose in the
    /// property-read path at the same time: the paths were not asking the same
    /// thing.
    ///
    /// The tree answers first, because it can express things a set cannot. It
    /// is canonical by construction, so the case juggling above becomes
    /// unnecessary; it is typed (`UserGlobalKind::Type`), so a same-named
    /// function cannot answer for a type; and it applies the real three tiers,
    /// so a user type reachable only through its enclosing `Namespace` or an
    /// `Imports` shadows correctly — which no `contains()` can say.
    ///
    /// `defined_classes` remains as a fallback, deliberately. Both are filled
    /// from the ONE call in `declare_class_identity`, so they agree on
    /// membership — but `declare_user_namespace_member` bails when a path
    /// segment is already occupied by a host leaf, which leaves a declaration
    /// in the set and not in the tree. Dropping the fallback would silently
    /// stop shadowing exactly those names. It goes when that registration gap
    /// closes (namespaceplan Phase 6), not before.
    pub(crate) fn shadows_builtin_type(&self, name: &str) -> bool {
        if self.resolve_user_namespace_type(name).is_some() {
            return true;
        }
        self.defined_classes.contains(name)
            || self.defined_classes.contains(&self.canon(name))
            || (!self.case_sensitive
                && self
                    .defined_classes
                    .iter()
                    .any(|g| g.eq_ignore_ascii_case(name)))
    }

    /// The shadow question, folded the way the NAMESPACE-TREE lookups fold.
    ///
    /// `shadows_builtin_type` honours the language's case policy — correct
    /// wherever the guarded lookup does too. But the platform registry
    /// (`lookup_type_ctor_spec`, `lookup_type_property_*`,
    /// `canonical_control_name`) matches case-INSENSITIVELY, and several
    /// callers reach it with a `normalize_type_hint`-lowercased spelling. In a
    /// case-sensitive language `shadows_builtin_type("label")` then cannot see
    /// the user's `class Label`, while the registry still answers — the guard
    /// and the lookup disagree about one name, and the user's class loses its
    /// own field writes to the DOM. A guard must fold exactly like the lookup
    /// it guards. ⛔ The lookups it guards are NO LONGER case-blind — the tree
    /// folds only where the language declares a fold. This stays case-blind on
    /// purpose: a user class shadows a platform type of the same name whatever
    /// the language's case rules, and that is a question about the USER's
    /// declaration, not about the tree's.
    pub(crate) fn user_owns_type_spelling(&self, name: &str) -> bool {
        self.shadows_builtin_type(name)
            || self
                .defined_classes
                .iter()
                .any(|g| g.eq_ignore_ascii_case(name))
    }

    /// Whether the receiver's own class DECLARES this member.
    ///
    /// A class owns its members: `Class Counter` with a `Count` property must
    /// read ITS `Count`, never a generic collection length. The uppercase
    /// heuristic used to hide this by accident — VB spells locals and `Me` in
    /// PascalCase, so capitalised receivers were excluded wholesale and the
    /// declared member survived. Asking the real question covers the lowercase
    /// receivers that heuristic never protected.
    pub(crate) fn receiver_class_declares_member(
        &self,
        receiver: &Expression,
        field: &str,
    ) -> bool {
        let Some(type_hint) =
            crate::primitives::calls::resolve_receiver_type_hint(self, receiver)
                .or_else(|| self.infer_expr_type_hint(receiver))
        else {
            return false;
        };
        let Some(class_name) = self.resolve_pending_class_name_for_type_hint(&type_hint) else {
            return false;
        };
        // ⛔ ONE ANSWER: `Compiler::resolution_chain` — a hand-rolled `parent`
        // climb is blind to every base but the first.
        for name in self.resolution_chain(&class_name) {
            let Some(pending) = self.pending_classes.get(name.as_str()) else {
                continue;
            };
            let key = self.js_member_storage_name_for_class(&name, field);
            // ⛔ A PROPERTY is declared under NEITHER name. `Public Property
            // Count` lands as the backing FIELD `_count` plus the accessor
            // `__set_count` — so asking for `count` in both maps misses it, and
            // the class's own property lost to the generic collection length.
            // `--dump-classes` shows exactly that: `fields: _count`,
            // `members: __set_count`, and no `count` anywhere.
            let setter = format!("__set_{key}");
            let getter = format!("__get_{key}");
            // ⛔ STATIC MEMBERS COUNT AS DECLARED TOO. This asked
            // `instance_field_types` and `instance_member_names` and nothing
            // else, so a `Shared`/`static` member was invisible here — and the
            // two spellings this guard exists to protect are `Count` and
            // `Length`:
            //
            //     Class F
            //         Public Shared Count As Integer = 5
            //     End Class          ' F.Count read as a COLLECTION LENGTH
            //
            // An INSTANCE field named `Count` was protected and answered 5; the
            // `Shared` one fell through to the generic length and answered
            // nothing at all. Renaming it to `Total` fixed it, which is the
            // signature of a name table beating a declaration.
            //
            // Worse than a plain miss: excluding a type receiver from
            // `is_size_property_read` pushes it into `size_member_may_be_
            // callable` instead of out of both, so the miss lands in the
            // read-then-call-if-callable path rather than on the ordinary
            // field read.
            if pending.instance_field_types.contains_key(&key)
                || pending.static_fields.iter().any(|n| n == &key)
                || pending
                    .instance_member_names
                    .iter()
                    .any(|n| n == &key || n == &setter || n == &getter)
            {
                return true;
            }
        }
        false
    }

    /// Whether a receiver expression NAMES A TYPE rather than denoting a value.
    ///
    /// `Foo.Count` where `Foo` is a class is a STATIC access; `foo.Count` on an
    /// instance is a member read. The two need different code, and the question
    /// is "does this identifier resolve to a type" — asked here as a real
    /// lookup.
    ///
    /// ⛔ This replaces `name.chars().next().is_ascii_uppercase()`, which used
    /// PascalCase as a proxy for "is a type". That is a NAMING CONVENTION doing
    /// a semantic job — a language check in disguise, in the same shape as the
    /// `!case_sensitive` ones. It is wrong in both directions: a lowercase type
    /// name is not recognised, and a local named `Total` is mistaken for a type.
    ///
    /// A binding in scope wins: a local SHADOWS a type of the same spelling, so
    /// a name that resolves to a slot is a value no matter what else shares it.
    pub(crate) fn receiver_names_a_type(&self, receiver: &Expression) -> bool {
        let ExprKind::Ident(name) = &receiver.kind else {
            return false;
        };
        if self.scope().resolve(name).is_some() {
            return false;
        }
        self.user_owns_type_spelling(name)
            || vybe_runtime::namespaces::is_registered_type(
                &self.profile.namespaces.type_scopes,
                &Self::tree_type_key(name),
                self.tree_fold(),
            )
    }

    pub(super) fn normalize_type_hint(type_hint: &str) -> String {
        type_hint.trim().to_lowercase()
    }

    /// The same type hint, keyed for a NAMESPACE TREE lookup.
    ///
    /// ⛔ Differs from [`Self::normalize_type_hint`] in exactly one way: it does
    /// NOT fold case. The tree stores every type under the spelling its
    /// registrar declared (`List`, `StringBuilder`, `LocalDate`), and since the
    /// fold became conditional a case-sensitive language asks the tree for an
    /// EXACT match. Handing a tree lookup the folded key makes
    /// `seg_eq("List", "list", None)` false, which is invisible in every
    /// folding language and takes out every lookup in the others.
    ///
    /// `normalize_type_hint` stays folded because its other consumers compare
    /// it against lowercase literals (`"integer"`, `"string"`). The two
    /// questions are different, so they get different functions rather than one
    /// function with a flag.
    pub(super) fn tree_type_key(type_hint: &str) -> String {
        // ⛔ Generic ARGUMENTS are not part of the key. A platform registers
        // `HashMap`, never `HashMap<String, Integer>`, so a hint that still
        // carries its arguments finds nothing. `resolve_source_type_alias`
        // erases them on the paths that go through it; a hint read straight
        // out of a declaration has not been through it yet.
        crate::primitives::generics::generic_base_name(type_hint.trim()).to_string()
    }

    pub(super) fn emit_default_value_for_type_hint(&mut self, type_hint: Option<&str>) {
        match type_hint.map(Self::normalize_type_hint).as_deref() {
            Some("integer") | Some("int") | Some("int32") | Some("longint") | Some("real")
            | Some("double") | Some("float") | Some("single") | Some("decimal") | Some("long")
            | Some("int64") | Some("short") | Some("int16") | Some("uint") | Some("uint32")
            | Some("ulong") | Some("uint64") | Some("ushort") | Some("uint16") | Some("byte")
            | Some("sbyte") => inst!(self, core_wasm::f64_const, 0.0),
            // Same question as the coercion arm in `arrays.rs`: a `char` that
            // holds a character defaults to `""`, one that is an 8-bit integer
            // defaults to 0. The language answers via `[builtin_types]`.
            Some("char") if self.hint_is_builtin_string("char") => {
                self.emit_const(Value::String(Arc::from("")))
            }
            Some("char") => inst!(self, core_wasm::f64_const, 0.0),
            Some("boolean") | Some("bool") => inst!(self, core_wasm::bool_const, false),
            Some(type_hint) if Self::is_string_type_hint(type_hint) => {
                self.emit_const(Value::String(Arc::from("")))
            }
            _ => self.emit_null(),
        }
    }

    /// Whether `type_hint` names a string.
    ///
    /// The spellings moved to `vybe_ast::builtin_types` (`builtinslotplan.md`
    /// step 4) so a profile can extend them; this stays as the shape its call
    /// sites want. Same answers — `builtin_types::tests` transcribes the old
    /// body's list and asserts every entry still classifies.
    ///
    /// Takes no `&self` deliberately: making it profile-aware here would change
    /// behaviour inside a move. `Compiler::builtin_type_of` is the
    /// profile-aware entry point.
    pub(super) fn is_string_type_hint(type_hint: &str) -> bool {
        vybe_ast::builtin_types::is(type_hint, vybe_ast::builtin_slots::BuiltinType::String)
    }

    /// Whether `type_hint` names any numeric type.
    ///
    /// The table now records WHICH numeric — see
    /// `builtin_types::PLATFORM_SPELLINGS` — which is what step 3 recorded as
    /// unresolvable. This predicate keeps collapsing that back to a bool for
    /// the call sites that only ask "is it a number".
    pub(super) fn is_numeric_type_hint(type_hint: &str) -> bool {
        vybe_ast::builtin_types::is_numeric(type_hint)
    }

    pub(super) fn fortran_out_param_ctor_name(type_hint: &str) -> Option<String> {
        let normalized = Self::normalize_type_hint(type_hint);
        if normalized.ends_with("()")
            || Self::is_numeric_type_hint(&normalized)
            || Self::is_string_type_hint(&normalized)
            || matches!(normalized.as_str(), "boolean" | "bool")
        {
            return None;
        }

        if let Some(inner) = normalized
            .strip_prefix("type(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            return Some(inner.trim().to_string());
        }

        if let Some(inner) = normalized
            .strip_prefix("class(")
            .and_then(|inner| inner.strip_suffix(')'))
        {
            return Some(inner.trim().to_string());
        }

        Some(normalized)
    }

    pub(super) fn maybe_initialize_fortran_out_param(&mut self, param: &Param) {
        if !self.profile.out_params_default_initialized || param.pass_by != PassBy::Out {
            return;
        }

        let Some(type_hint) = param.type_hint.as_deref() else {
            return;
        };
        let Some(ctor_name) = Self::fortran_out_param_ctor_name(type_hint) else {
            return;
        };
        let Some(slot) = self.scope().resolve(&param.name) else {
            return;
        };
        if !(self.defined_classes.contains(&ctor_name)
            || self.defined_globals.contains(&ctor_name)
            || self.profile.lookup_known_type(&ctor_name).is_some())
        {
            return;
        }

        self.emit_u16(Op::LOCAL_GET, slot);
        self.emit(Op::REF_IS_NULL);
        let line = self.line;
        self.chunk().emit_if(line);

        if let Some((module, func)) = self
            .profile
            .lookup_known_type(&ctor_name)
            .map(|(module, func)| (module.to_string(), func.to_string()))
        {
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, 0);
        } else {
            self.emit_global_read(&ctor_name);
            self.emit_direct_callable_invoke(0);
        }
        self.emit_u16(Op::LOCAL_SET, slot);

        self.chunk().emit_end(line);
    }

    pub(super) fn can_instantiate_fortran_ctor_name(&self, ctor_name: &str) -> bool {
        self.defined_classes.contains(ctor_name)
            || self.defined_globals.contains(ctor_name)
            || self.profile.lookup_known_type(ctor_name).is_some()
    }

    pub(super) fn emit_fortran_ctor_call(&mut self, ctor_name: &str) {
        if let Some((module, func)) = self
            .profile
            .lookup_known_type(ctor_name)
            .map(|(module, func)| (module.to_string(), func.to_string()))
        {
            let idx = self.import(&module, &func);
            self.emit_host_call(idx, 0);
        } else {
            self.emit_global_read(ctor_name);
            self.emit_direct_callable_invoke(0);
        }
    }

    pub(super) fn fortran_allocate_ctor_name(&self, target: &Expression) -> Option<String> {
        let type_hint = self.infer_expr_type_hint(target)?;
        let normalized = Self::normalize_type_hint(&type_hint);
        let element_hint = normalized
            .strip_suffix("()")
            .unwrap_or(normalized.as_str())
            .trim();
        let ctor_name = Self::fortran_out_param_ctor_name(element_hint)?;
        self.can_instantiate_fortran_ctor_name(&ctor_name)
            .then_some(ctor_name)
    }

    pub(super) fn emit_fortran_allocated_array(
        &mut self,
        dim_slots: &[u16],
        ctor_name: Option<&str>,
    ) {
        let line = self.line;
        self.emit_u16(Op::LOCAL_GET, dim_slots[0]);
        common::collections::emit_new_with_length(&mut self.chunks, self.current, line);
        let array_slot = self.define_local("__fortran_alloc_array");
        self.emit_u16(Op::LOCAL_SET, array_slot);

        if dim_slots.len() == 1 && ctor_name.is_none() {
            self.emit_u16(Op::LOCAL_GET, array_slot);
            return;
        }

        let idx_slot = self.define_local("__fortran_alloc_idx");
        self.emit_const(Value::F64(0.0));
        self.emit_u16(Op::LOCAL_SET, idx_slot);

        let block_patch = self.chunk().emit_block(line);
        let (loop_patch, _) = self.chunk().emit_loop_s(line);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_u16(Op::LOCAL_GET, dim_slots[0]);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        };
        self.emit(Op::I32_EQZ);
        self.chunk().emit_br_if(1, line);

        self.emit_u16(Op::LOCAL_GET, array_slot);
        self.emit_u16(Op::LOCAL_GET, idx_slot);
        if dim_slots.len() > 1 {
            self.emit_fortran_allocated_array(&dim_slots[1..], ctor_name);
        } else if let Some(ctor_name) = ctor_name {
            self.emit_fortran_ctor_call(ctor_name);
        } else {
            self.emit_null();
        }
        common::collections::emit_set(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);

        self.emit_u16(Op::LOCAL_GET, idx_slot);
        self.emit_const(Value::F64(1.0));
        self.emit(Op::F64_ADD);
        self.emit_u16(Op::LOCAL_SET, idx_slot);
        self.chunk().emit_br(0, line);
        self.chunk().emit_end(line);
        self.chunk().patch_loop(loop_patch);
        self.chunk().emit_end(line);
        self.chunk().patch_block(block_patch);

        self.emit_u16(Op::LOCAL_GET, array_slot);
    }

    pub(super) fn expr_prefers_numeric_add(&self, expr: &Expression) -> bool {
        self.infer_expr_type_hint(expr)
            .as_deref()
            .is_some_and(Self::is_numeric_type_hint)
    }

    pub(super) fn compile_expr_with_numeric_add_hint(
        &mut self,
        expr: &Expression,
        prefer_numeric_add: bool,
    ) -> Result<(), String> {
        if prefer_numeric_add {
            if let ExprKind::Binary {
                op: BinOp::Add,
                left,
                right,
            } = &expr.kind
            {
                self.compile_expr_with_numeric_add_hint(left, true)?;
                self.compile_expr_with_numeric_add_hint(right, true)?;
                self.emit(Op::F64_ADD);
                return Ok(());
            }
        }

        self.compile_expr(expr)
    }

    pub(super) fn emit_assignment_type_coercion_for_target(&mut self, target: &Expression) {
        let ExprKind::Ident(name) = &target.kind else {
            return;
        };
        self.emit_assignment_type_coercion_for_ident(name);
    }

    pub(super) fn emit_assignment_type_coercion_for_ident(&mut self, name: &str) {
        if self.lookup_array_binding(name).is_some() {
            return;
        }
        let Some(type_hint) = self.lookup_var_type_hint(name).map(str::to_string) else {
            return;
        };
        let normalized = Self::normalize_type_hint(&type_hint);
        inst!(self, core_wasm::dup);
        self.emit(Op::REF_IS_NULL);
        self.emit(Op::I32_EQZ);
        let line = self.line;
        self.chunk().emit_if(line);
        match normalized.as_str() {
            "integer" | "int" | "int32" | "longint" | "long" | "int64" | "short" | "int16"
            | "uint" | "uint32" | "ulong" | "uint64" | "ushort" | "uint16" | "byte" | "sbyte" => {
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
                common::convert::emit_to_int(self.chunk(), line);
            }
            "real" | "double" | "float" | "single" | "decimal" => {
                let number_idx = self.import("ecma:number", "Number");
                self.emit_host_call(number_idx, 1);
            }
            _ => {}
        }
        self.chunk().emit_end(line);
    }

    pub(super) fn emit_file_key_compare(&mut self, relation: FileKeyRelation) {
        match relation {
            FileKeyRelation::Equal => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
            }
            FileKeyRelation::Greater => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
            }
            FileKeyRelation::GreaterOrEqual => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
            }
            FileKeyRelation::Less => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_lt(self.chunk(), line);
            }
            FileKeyRelation::LessOrEqual => {
                let line = self.line;
                crate::primitives::ops::emit_dyn_le(self.chunk(), line);
            }
        }
    }

    pub(super) fn emit_global_map_get_into_local(
        &mut self,
        map_name: &str,
        key_slot: u16,
        value_slot: u16,
    ) {
        self.emit_ensure_global_map(map_name);
        self.emit_global_read(map_name);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit(Op::ARRAY_GET);
        self.emit_u16(Op::LOCAL_SET, value_slot);
    }

    pub(super) fn emit_global_map_set_from_local(
        &mut self,
        map_name: &str,
        key_slot: u16,
        value_slot: u16,
    ) {
        self.emit_ensure_global_map(map_name);
        self.emit_global_read(map_name);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit(Op::ARRAY_SET);
    }

    pub(super) fn emit_global_map_set_const(
        &mut self,
        map_name: &str,
        key_slot: u16,
        value: Value,
    ) {
        self.emit_ensure_global_map(map_name);
        self.emit_global_read(map_name);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_const(value);
        self.emit(Op::ARRAY_SET);
    }

    pub(super) fn emit_global_map_set_null(&mut self, map_name: &str, key_slot: u16) {
        self.emit_ensure_global_map(map_name);
        self.emit_global_read(map_name);
        self.emit_u16(Op::LOCAL_GET, key_slot);
        self.emit_null();
        self.emit(Op::ARRAY_SET);
    }

    pub(super) fn emit_record_rows_cache(&mut self, file_slot: u16, rows_slot: u16, len_slot: u16) {
        let line = self.line;

        self.emit_global_map_get_into_local("__vb_record_rows_by_handle", file_slot, rows_slot);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        self.emit(Op::REF_IS_NULL);
        self.chunk().emit_if(line);

        self.emit_ensure_global_map("__vb_file_path_by_handle");
        self.emit_global_read("__vb_file_path_by_handle");
        self.emit_u16(Op::LOCAL_GET, file_slot);
        self.emit(Op::ARRAY_GET);
        crate::primitives::fs_path::emit_read_file(self.chunk(), line);
        self.emit_const(Value::String(Arc::from("\n")));
        fn_call!(self, "ecma:string", "split", 2);
        self.emit_u16(Op::LOCAL_SET, rows_slot);

        let skip_trim = self.chunk().emit_block(line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_array_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        inst!(self, core_wasm::i32_const, 0);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(0, line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        self.emit_u16(Op::LOCAL_GET, len_slot);
        inst!(self, core_wasm::i32_const, 1);
        self.emit(Op::I32_SUB);
        common::collections::emit_get(&mut self.chunks, self.current, line);
        self.emit_const(Value::String(Arc::from("")));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_eq(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_not(self.chunk(), line);
        };
        self.chunk().emit_br_if(0, line);
        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_pop(&mut self.chunks, self.current, line);
        self.emit(Op::DROP);
        self.chunk().emit_end(line);
        self.chunk().patch_block(skip_trim);

        self.emit_global_map_set_from_local("__vb_record_rows_by_handle", file_slot, rows_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, rows_slot);
        common::collections::emit_array_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, len_slot);
    }

    pub(super) fn emit_record_assign_nulls(&mut self, variables: &[String]) {
        for variable in variables {
            self.emit_null();
            self.emit_var_set(variable);
        }
    }

    pub(super) fn emit_record_assign_values_from_local(
        &mut self,
        values_slot: u16,
        variables: &[String],
    ) {
        for (index, variable) in variables.iter().enumerate() {
            self.emit_u16(Op::LOCAL_GET, values_slot);
            self.emit_const(Value::F64(index as f64));
            self.emit(Op::ARRAY_GET);
            self.emit_assignment_type_coercion_for_ident(variable);
            self.emit_var_set(variable);
        }
    }

    pub(super) fn emit_record_rewrite_field_format(
        &mut self,
        field_format: Option<&RecordFieldFormat>,
    ) {
        let Some(field_format) = field_format else {
            return;
        };

        let number_idx = self.import("ecma:number", "Number");
        let to_fixed_idx = self.import("ecma:number", "toFixed");
        self.emit_host_call(number_idx, 1);
        self.emit_const(Value::F64(field_format.decimal_places as f64));
        self.emit_host_call(to_fixed_idx, 2);
    }

    pub(super) fn vb_fixed_string_len(type_hint: &str) -> Option<i32> {
        let normalized = Self::normalize_type_hint(type_hint);
        let (base, len) = normalized.split_once('*')?;
        let base = base.trim();
        if base != "string" && base != "system.string" && !base.ends_with(".string") {
            return None;
        }
        len.trim().parse::<i32>().ok().filter(|len| *len >= 0)
    }

    pub(super) fn emit_vb_fixed_string_adjust_from_stack(
        &mut self,
        target_len: i32,
        align_right: bool,
    ) {
        let line = self.line;
        let value_slot = self.define_local("__vb_fixed_string_value");
        let to_string = self.import("ecma:string", "String");
        let pad_idx = self.import(
            "ecma:string",
            if align_right { "padStart" } else { "padEnd" },
        );

        self.emit_host_call(to_string, 1);
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_const(Value::I32(target_len));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_gt(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_const(Value::I32(target_len));
        common::strings::emit_substring(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, value_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, value_slot);
        self.emit_const(Value::I32(target_len));
        self.emit_const(Value::String(Arc::from(" ")));
        self.emit_host_call(pad_idx, 3);
    }

    pub(super) fn compile_vb_fixed_string_stmt(
        &mut self,
        target: &Expression,
        value: &Expression,
        align_right: bool,
    ) -> Result<(), String> {
        let ExprKind::Ident(name) = &target.kind else {
            self.compile_expr(value)?;
            self.emit(Op::DROP);
            return Ok(());
        };
        let Some(type_hint) = self.lookup_var_type_hint(name) else {
            self.compile_expr(value)?;
            self.compile_assign_target(target)?;
            return Ok(());
        };
        let Some(target_len) = Self::vb_fixed_string_len(type_hint) else {
            self.compile_expr(value)?;
            self.compile_assign_target(target)?;
            return Ok(());
        };

        self.compile_expr(value)?;
        self.emit_vb_fixed_string_adjust_from_stack(target_len, align_right);
        self.compile_assign_target(target)
    }

    pub(super) fn compile_vb_mid_stmt(
        &mut self,
        target: &Expression,
        start: &Expression,
        count: &Expression,
        value: &Expression,
    ) -> Result<(), String> {
        let line = self.line;
        let target_slot = self.define_local("__vb_mid_target");
        let start_slot = self.define_local("__vb_mid_start");
        let count_slot = self.define_local("__vb_mid_count");
        let value_slot = self.define_local("__vb_mid_value");
        let prefix_slot = self.define_local("__vb_mid_prefix");
        let replace_slot = self.define_local("__vb_mid_replace");
        let to_string = self.import("ecma:string", "String");

        self.compile_expr(target)?;
        self.emit_u16(Op::LOCAL_SET, target_slot);
        self.compile_expr(start)?;
        common::convert::emit_to_int(self.chunk(), line);
        self.emit_const(Value::I32(1));
        self.emit(Op::I32_SUB);
        self.emit_u16(Op::LOCAL_SET, start_slot);

        self.emit_u16(Op::LOCAL_GET, start_slot);
        inst!(self, core_wasm::i32_const, 0);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.emit(Op::I32_EQZ);
        self.chunk().emit_if(line);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_u16(Op::LOCAL_SET, start_slot);
        self.chunk().emit_end(line);

        self.compile_expr(value)?;
        self.emit_host_call(to_string, 1);
        self.emit_u16(Op::LOCAL_SET, value_slot);

        self.compile_expr(count)?;
        common::convert::emit_to_int(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, count_slot);

        self.emit_u16(Op::LOCAL_GET, count_slot);
        self.emit_const(Value::I32(0));
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_ge(self.chunk(), line);
        };
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_to_bool(self.chunk(), line);
        }
        self.chunk().emit_if(line);
        self.emit_u16(Op::LOCAL_GET, count_slot);
        self.emit_u16(Op::LOCAL_SET, replace_slot);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        common::strings::emit_length(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, replace_slot);
        self.chunk().emit_end(line);

        self.emit_u16(Op::LOCAL_GET, target_slot);
        inst!(self, core_wasm::i32_const, 0);
        self.emit_u16(Op::LOCAL_GET, start_slot);
        common::strings::emit_substring(self.chunk(), line);
        self.emit_u16(Op::LOCAL_SET, prefix_slot);

        self.emit_u16(Op::LOCAL_GET, prefix_slot);
        self.emit_u16(Op::LOCAL_GET, value_slot);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };
        self.emit_u16(Op::LOCAL_GET, target_slot);
        self.emit_u16(Op::LOCAL_GET, start_slot);
        self.emit_u16(Op::LOCAL_GET, replace_slot);
        self.emit(Op::I32_ADD);
        self.emit_u16(Op::LOCAL_GET, target_slot);
        common::strings::emit_length(self.chunk(), line);
        common::strings::emit_substring(self.chunk(), line);
        {
            let line = self.line;
            crate::primitives::ops::emit_dyn_add(self.chunk(), line);
        };

        if let ExprKind::Ident(name) = &target.kind {
            if let Some(type_hint) = self.lookup_var_type_hint(name) {
                if let Some(target_len) = Self::vb_fixed_string_len(type_hint) {
                    self.emit_vb_fixed_string_adjust_from_stack(target_len, false);
                }
            }
        }

        self.compile_assign_target(target)
    }

    pub(super) fn compile_vb_err_raise_stmt(&mut self, args: &[Argument]) -> Result<(), String> {
        if let Some(description) = args.get(2).or_else(|| args.get(1)).or_else(|| args.first()) {
            self.compile_expr(&description.value)?;
        } else {
            self.emit_const(Value::String(Arc::from("")));
        }

        self.emit_js_exception_ctor_from_message_value("Exception")?;

        if let Some(number) = args.first() {
            inst!(self, core_wasm::dup);
            self.compile_expr(&number.value)?;
            let key = self.resolve_slot_interned(&class_slots::ClassSlot::internal("number"));
            self.class_set_resolved(
                class_slots::ObjSource::Stack,
                &key,
                class_slots::ValueSource::Stack,
            );
        }

        if let Some(source) = args.get(1) {
            inst!(self, core_wasm::dup);
            self.compile_expr(&source.value)?;
            self.class_set(
                class_slots::ObjSource::Stack,
                &class_slots::ClassSlot::internal("source"),
                class_slots::ValueSource::Stack,
            );
        }

        let line = self.line;
        common::errors::emit_throw(self.chunk(), line);
        Ok(())
    }

    pub(super) fn is_collection_like_type_hint(type_hint: &str) -> bool {
        let normalized = Self::normalize_type_hint(type_hint);
        let bare = normalized
            .split('<')
            .next()
            .unwrap_or(normalized.as_str())
            .trim_end_matches('?');
        let terminal = bare.rsplit('.').next().unwrap_or(bare);
        Self::is_string_type_hint(type_hint)
            || matches!(
                terminal,
                "list"
                    | "arraylist"
                    | "dictionary"
                    | "queue"
                    | "stack"
                    | "hashset"
                    | "sortedset"
                    | "set"
                    | "collection"
                    | "icollection"
                    | "readonlycollection"
                    | "enumerable"
                    | "ienumerable"
                    | "readonlylist"
                    | "ilist"
                    | "array"
            )
            || bare.ends_with("[]")
            || normalized.ends_with("()")
    }
}
