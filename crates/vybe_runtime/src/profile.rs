//! Language profile — compilation semantics per language.
//!
//! The grammar defines syntax (how to parse). The profile defines semantics
//! (how to compile). Together they fully describe a language.
//!
//! Profiles are loaded from `languages/<lang>/profile` files — no hardcoded
//! language knowledge lives in Rust code.

use std::collections::HashMap;

/// Compilation semantics for a language.
#[derive(Debug, Clone)]
pub struct LanguageProfile {
    /// Profile short name — matches the `[info].name` TOML field
    /// (`"js"`, `"vb"`, `"csharp"`, `"python"`, …). Used by cross-
    /// language orchestration code in `common::*` to dispatch to the
    /// right per-language helper (e.g. `normalize_class`).
    pub name: String,

    /// How functions return values.
    pub function_return: ReturnStyle,

    /// Local slot name for the result of a function in `ResultSlot` mode.
    /// Pascal uses `"Result"` because user code writes `Result := X`.
    /// VB uses an internal name like `"__result__"` because user code writes
    /// `FunctionName := X` (matched via current_func_name in the compiler) —
    /// keeping the slot name internal lets users declare a class called
    /// `Result` without it being shadowed by the function's return slot.
    /// Defaults to `"Result"` for backward compatibility.
    pub result_slot_name: String,

    /// The keyword used for `self`/`this` in methods.
    pub self_keyword: String,

    /// The keyword for base/super class reference (e.g. "mybase", "super", "inherited").
    pub base_keyword: Option<String>,

    /// Constructor method name (matched case-insensitively for case-insensitive languages).
    pub constructor_name: String,

    /// Exception type thrown by the common emitter when a method is called on
    /// a `null`/`undefined` receiver. Language-defined and cross-language
    /// compatible: JS throws `TypeError`, PHP throws `Error`, etc. The emitter
    /// stays language-agnostic — it just throws whatever the profile names.
    pub member_call_on_null_error: String,

    /// Class instance-method dispatch model.
    /// "instance" (default): construction binds compiled method refs
    /// directly onto the instance.
    /// "prototype": the class carries its instance methods on an open
    /// method-table object (`prototype`) and construction binds from it,
    /// so post-definition reassignment (`C.prototype.m = wrap(...)`)
    /// reaches instances constructed afterwards (ECMA-262 §15.7, Python
    /// `__dict__`-style open classes).
    pub class_method_dispatch: String,

    /// Whether enum values are compiled as global ordinal constants.
    pub enum_as_ordinals: bool,

    /// Whether the language is case-sensitive.
    pub case_sensitive: bool,

    /// String indexing: "zero_based" or "one_based" (VB).
    pub string_indexing: StringIndexing,

    /// Whether array upper bounds are inclusive (VB: Dim arr(5) = 6 elements).
    pub array_upper_bound_inclusive: bool,

    /// Negative array indices wrap from the end (Python `arr[-1]`,
    /// Ruby, PHP, Dart `.last`-style). JS / VB / C# return undefined
    /// for negative — keep this `false` for them; ECMA-262 §10.4.2.1
    /// is the JS-conformant default.
    pub negative_index_wraps: bool,

    /// A zero slice step raises the language value-error (Python
    /// `ValueError: slice step cannot be zero`). Languages with lenient
    /// slicing (or without strided slices) keep this `false` → empty result.
    pub slice_step_zero_raises: bool,

    /// Tuple literals produce a distinct tuple value (array-backed, but tagged
    /// via `vybe_compiler::primitives::tuples::TUPLE_TAG`) rather than a plain list, so
    /// `repr`/`type()`/slicing tell a tuple from a list. Python opts in;
    /// languages whose `(a, b)` is just grouping/an array keep this `false`.
    pub tuple_literals_tagged: bool,

    /// Whether parens are used for both calls and indexing (VB: arr(i)).
    pub parens_for_index: bool,
    /// The language has ONE array type covering both list and dictionary use
    /// (PHP `array`), so an indexed write must decide the backing
    /// representation at runtime: a string key promotes an empty sequential
    /// array to an ordered Map, and `$x[$k][] = $v` auto-vivifies the missing
    /// inner array rather than faulting.
    ///
    /// Languages with separate list and map types know the representation
    /// statically and need none of this.
    pub unified_array_map: bool,

    /// The string-concatenation operator coerces BOTH operands to string
    /// before joining, rather than leaning on the concat op's own coercion —
    /// PHP `.` and Lua `..`. The spelling of that coercion can differ from the
    /// shared one, in which case the language also registers a
    /// `LanguageHooks::concat_stringify`.
    pub concat_stringifies_operands: bool,

    /// Entry point function name to auto-call if defined (e.g. "main").
    pub entry_point: Option<String>,

    /// JS: `var` declarations are hoisted to function scope.
    pub hoist_var: bool,

    /// JS: `+` operator uses dynamic add (string concat if either operand is string).
    pub dynamic_add: bool,

    /// The language has first-class function references (WASM `ref.func` /
    /// funcref). When set, referencing a function/static-method as a VALUE
    /// (not calling it) tears it off into a `REF_FUNC` funcref instead of
    /// invoking it or reading a bound-method property. Only WASM-level
    /// frontends (wast) that model real funcrefs enable this.
    pub function_references: bool,

    /// Arithmetic operands may be of a type not known until runtime
    /// (dynamically-typed languages: JS, PHP, Python, Ruby). When set,
    /// `-`/`*`/`/`/`%` whose operand types aren't statically resolved emit
    /// the runtime-polymorphic `emit_dyn_*` sequences (which dispatch
    /// BigInt → `i64.*`, else Number → `f64.*`). Statically-typed languages
    /// leave this off and keep direct typed opcodes.
    pub dynamic_numeric_dispatch: bool,

    /// JS: support `require()` for CommonJS module loading.
    pub commonjs_require: bool,

    /// Python: a function whose every `return` is a same-arity tuple
    /// literal compiles to WASM multi-value — callee pushes N values,
    /// caller destructures directly off the stack. Avoids the heap
    /// tuple allocation for the common `return a, b` / `a, b = f()` idiom.
    pub multi_value_tuple_returns: bool,

    /// Reading a method off an instance yields a fresh bound-method object,
    /// distinct per instance and carrying its own receiver (Python/Ruby
    /// descriptor semantics: `C().f is C().f` is False). When false, the
    /// language binds the receiver at call time and instance method reads share
    /// the underlying function (JS: `a.m === b.m` is True). General behavior,
    /// gated on the profile — never a language-name check.
    pub methods_bind_on_access: bool,

    /// A parameter's default expression is evaluated ONCE at definition and
    /// reused on every call that omits the arg (Python/Ruby: the classic
    /// mutable-default — `def f(a=[])` shares one list; `f() is f()` is True).
    /// When false (JS/C#/… per spec) the default is re-evaluated per call.
    /// General behavior, gated on the profile — never a language-name check.
    pub default_args_evaluated_once: bool,

    /// VB: ByRef args wrapped in single-element arrays for call-by-reference.
    pub byref_boxing: bool,

    /// VB: With obj ... End With — bare .Member resolves to With target.
    pub with_block: bool,

    /// VB: New Foo() With { .Prop = val } initializer syntax.
    pub new_with_initializer: bool,

    /// VB: New List(Of T) From { items } initializer syntax.
    pub new_from_initializer: bool,

    /// C/JS: switch statements fall through to the next case unless there's
    /// an explicit `break`. VB/Pascal/Python: each case is independent.
    pub switch_fallthrough: bool,

    /// The language has ECMAScript-style private class members (`#field`).
    /// Only languages that declare this treat a `#`-prefixed member specially;
    /// the shared compiler no longer keys that on the JS name.
    pub supports_private_fields: bool,

    /// Functions/class-constructors are first-class objects carrying
    /// `Function.prototype` methods `.bind`/`.call`/`.apply` (ECMAScript
    /// §20.2.3). Languages that declare this route those member calls through
    /// the function-object path instead of instance-method dispatch.
    pub has_function_prototype_bind: bool,

    /// `.call(…)` / `.apply(…)` written on an arbitrary receiver mean
    /// "invoke this callable", not "call the member of that name". Languages
    /// that declare this route such calls to the function-object path when the
    /// receiver has no own method of that name.
    ///
    /// Off by default, because in most languages `apply`/`call` are perfectly
    /// ordinary method names: routing them to the function builtins made a
    /// user method named `apply` return null, and one named `call` panic the
    /// host (`ecma` function slicing `args[2..]`) — in Dart, PHP, Python and
    /// JS alike. This is a PROPERTY precisely so the behaviour is declared by
    /// each language rather than carved out by name in shared code.
    pub function_invocation_members: bool,

    /// The global `Function` is a constructor that builds a function from
    /// string arguments (ECMAScript §20.2.1.1). JS-only today.
    pub has_function_constructor: bool,

    /// An `async` function body is wrapped so a synchronous `throw` inside it
    /// becomes a rejected promise instead of propagating (ECMAScript §27.7).
    /// Languages without promise-based async leave this off.
    pub async_wraps_body_in_try: bool,

    /// Exception objects are shaped as ECMAScript `Error`s: internal fields
    /// (`message`/`name`/`__type`/`stack`) are non-enumerable (§20.5) and the
    /// object carries the `Error`/`TypeError`/… `instanceof` prototype chain.
    pub ecma_error_object_shape: bool,

    /// The language has a distinct `undefined` value separate from `null`
    /// (ECMAScript). When set, sites that need "the empty/default value"
    /// — an unmatched `find`, an initializer-less declared field — emit
    /// `undefined`; otherwise they emit `null` (Python `None`, VB `Nothing`,
    /// .NET/PHP `null`).
    pub has_undefined_value: bool,

    /// Spread arguments (`f(...arr)`, `arr.push(...xs)`) are expanded at the
    /// call site by the shared compiler (ECMAScript spread). Languages whose
    /// walker lowers their own spread syntax to `Argument { spread: true }`
    /// opt in so the same expansion path serves them.
    pub supports_spread_arguments: bool,

    /// The language has ECMAScript module runtime features — dynamic
    /// `import()` (lowered by the walker to a `__js_dynamic_import` call) and
    /// the `import.meta` object. JS-only today.
    pub supports_dynamic_import: bool,

    /// The language has an ECMAScript `BigInt` type backed by the `ecma:bigint`
    /// host surface, so arbitrary-precision operators (e.g. `**`) route there.
    pub has_ecma_bigint: bool,

    /// The language honors ECMAScript strict mode — a top-level `"use strict"`
    /// prologue sets strict semantics (e.g. assignment to an undeclared global
    /// throws, §11.2.1). Off for languages with no strict-mode concept.
    pub ecma_strict_mode: bool,

    /// `switch` case matching uses ECMAScript strict equality (`===`, no type
    /// coercion, §14.12.1). Off for languages whose switch uses loose equality.
    pub ecma_switch_strict_equality: bool,

    /// Variable declarations follow ECMAScript lexical rules: `const` bindings
    /// are immutable (reassignment guarded), top-level `var` also binds on
    /// `globalThis`, and `const`/`let` init infers the function `.name`.
    pub ecma_lexical_declarations: bool,

    /// The language exposes the ECMAScript global object surface — `Object`,
    /// `Array`, `Math`, `JSON`, built-in constructors, `Object.groupBy`, the
    /// `len`/`size` canonical mappings — recognized as runtime globals.
    pub has_ecma_globals: bool,

    /// Operators apply ECMAScript coercion: `+`/relational operators run
    /// `ToPrimitive`/`ToNumber` on their operands, `===` is strict equality,
    /// `>>>` yields an unsigned 32-bit Number, and compound assignments coerce
    /// likewise. Off for languages with their own operator typing.
    pub ecma_operator_coercion: bool,

    /// `==`/`!=` use abstract (loose) equality with cross-type coercion via the
    /// host `ecma:value.abstractEq` (ECMAScript §7.2.15; PHP `==` shares this).
    /// Off for languages whose `==` is strict/typed.
    pub abstract_equality: bool,

    /// Functions expose the ECMAScript `arguments` object binding the actual
    /// call arguments (ECMA-262 §10.4.4). Off for languages without it.
    pub has_arguments_object: bool,

    /// Calls tolerate arity mismatch the ECMAScript way — missing parameters
    /// bind to `undefined`, extra arguments are ignored (constructors included)
    /// rather than being an error. Off for languages that enforce arity.
    pub relaxed_call_arity: bool,

    /// `new` follows the ECMAScript `[[Construct]]` model: any function is a
    /// constructor, built-in constructors (`Set`/`Map`/…) dispatch through the
    /// host, rest parameters are packed, and `new.target` is bound. Off for
    /// languages whose construction is a plain class-instantiation.
    pub ecma_new_dispatch: bool,

    /// Array literals may have elisions (holes), e.g. `[1, , 3]` — the walker
    /// marks a hole with a sentinel key and the compiler builds a sparse array.
    /// ECMAScript-only syntax.
    pub ecma_array_elisions: bool,

    /// The language has generator functions the compiler can detect statically
    /// (JS, PHP) — a direct call to a generator yields an iterator without a
    /// runtime `isGenerator` check.
    pub has_generators: bool,

    /// Object literals follow ECMAScript semantics: insertion-ordered key
    /// tracking (`ecma:object.trackKey`), method shorthand compiled as
    /// functions, and `fn.name` inference from the property key. Off for
    /// languages whose map/record literals don't carry these.
    pub ecma_object_literals: bool,

    /// Logical/comparison operators yield real boolean values in expression
    /// position (ECMAScript). When set, results of `!x` and friends are
    /// materialized as `Bool` (see also `materialize_bool_results`, which Go
    /// uses for the same effect on comparisons).
    pub ecma_boolean_operators: bool,

    /// Property/element access is dynamic (ECMAScript): reads go through the
    /// dynamic `STRUCT_GET`/host lookup path, honoring optional chaining
    /// (`?.`, `?.[]`) and dynamic property names, rather than static typed
    /// field access. Off for statically-typed field layouts (VB/C#/Pascal).
    pub dynamic_member_access: bool,

    /// The language has the ECMAScript `typeof` operator, whose result strings
    /// (`"function"`, `"undefined"`, …) and never-throws semantics the compiler
    /// reproduces (ECMA-262 §13.5.3). Off for languages with no such operator.
    pub ecma_typeof_operator: bool,

    /// Numeric/string coercion of objects runs the ECMAScript `ToPrimitive`
    /// abstract operation (`Symbol.toPrimitive`/`valueOf`/`toString`,
    /// ECMA-262 §7.1.1) before the final conversion. Off for languages that
    /// coerce objects directly.
    pub ecma_to_primitive: bool,

    /// The language has the ECMAScript `in` operator testing property/private-
    /// field existence on an object (`"k" in obj`, `#f in obj`). Off for
    /// languages whose `in` (Python membership) is lowered to a call instead.
    pub ecma_in_operator: bool,

    /// Array instance methods (`map`/`filter`/`find`/`sort`/…) dispatch through
    /// the ECMAScript `ecma:array` host surface, and callback arguments are
    /// invoked with runtime dispatch. When false the language resolves array
    /// methods through its profile's array-method table instead.
    pub ecma_array_method_dispatch: bool,

    /// Promises expose ECMAScript chaining methods (`.then`/`.catch`/`.finally`,
    /// ECMA-262 §27.2.5) that the shared compiler lowers to a promise-chain
    /// call. Languages without ECMA promises leave this off.
    pub ecma_promise_methods: bool,

    /// Iterator `.next()` returns an ECMAScript result record `{ value, done }`
    /// (ECMA-262 §27.1.1.3), so the compiler shapes iteration results with an
    /// explicit `done` flag and `value` field. Off for languages whose
    /// iteration protocol signals completion differently.
    pub ecma_iterator_result_shape: bool,

    /// PHP/Java: `Throwable` is the universal exception root; `Exception`
    /// is only one branch (the `Error` branch is a sibling), so
    /// `catch (Exception)` must NOT be a catch-all — it matches via the
    /// `__types` inheritance chain instead. When false (Python/.NET/Ruby),
    /// `Exception` is the root and `catch (Exception)` catches everything.
    pub throwable_is_root: bool,

    /// Fields are resolved by the reference's DECLARED (static) type, so a
    /// subclass field of the same name HIDES the parent's rather than
    /// overriding it — both occupy distinct storage slots (java/C#/VB, C++).
    /// When false (JS/Python/PHP/Ruby: dynamic single-slot objects), a
    /// same-named subclass field OVERRIDES the parent's one slot.
    pub field_hiding: bool,

    /// Instance methods dispatch on the RUNTIME type by default, without any
    /// `virtual`/`Overridable` keyword (java/python/js/php/ruby/dart). When
    /// false (C#/VB/Pascal/C++), only a method the source explicitly marked
    /// `is_virtual`/`is_override`/`is_abstract` dispatches dynamically; every
    /// other method binds to the reference's DECLARED type — the same
    /// static-type rule [`field_hiding`] applies to fields, and what makes C#
    /// `new`-hiding work. The member-level marker for that is `is_hiding`
    /// (C# `new`, VB `Shadows`, Pascal `reintroduce`) — a DIFFERENT flag from
    /// `is_not_overridable` below, which they were conflated with.
    ///
    /// Languages that opt in still exclude `is_static` and
    /// `is_not_overridable` members (VB `NotOverridable`, java `final`), which
    /// can never be overridden and so keep their direct bind.
    pub methods_virtual_by_default: bool,

    /// Properties and methods occupy SEPARATE member namespaces (PHP:
    /// `$o->foo` the property and `$o->foo()` the method coexist). A property
    /// whose name collides with a method is stored under a mangled slot so
    /// the two don't clobber each other in the shared object model. When
    /// false (JS/…: one member namespace), no mangling.
    pub separate_property_method_namespace: bool,

    /// How reflection/runtime type identity names are formed. Each language
    /// owns its type namespace — this must NOT be hardcoded to one language's
    /// scheme. `Native` (default) keeps a type's own name (`Throwable`,
    /// `Error`), so a language's real hierarchy — and its exception `__types`
    /// chain — is preserved. `Dotnet` qualifies under `System.` and applies
    /// .NET BCL primitive naming (`int`→`Int32`), for C#/VB reflection.
    /// Historically this was hardcoded to `Dotnet` for every language, which
    /// force-flattened e.g. PHP's sibling `Error`/`Exception` into .NET's
    /// single `System.Exception` root — the bug this replaces.
    pub reflection_type_naming: ReflectionTypeNaming,

    /// PHP: relational operators (`<`/`>`/`<=`/`>=`/`<=>`) compare two
    /// strings lexicographically and otherwise fall back to numeric/dynamic
    /// comparison (DateTime operands are unboxed first). When false, the
    /// generic dynamic comparison is used.
    pub string_aware_relational: bool,

    /// Some frontends want boolean-valued operators to materialize as actual
    /// Bool values in expression position, not raw WASM i32 conditions. Go
    /// needs this so `fmt.Println(5 == 5)` prints `true` while
    /// `fmt.Println(1)` still prints `1`.
    pub materialize_bool_results: bool,

    /// An OBJECT can be called like a function: Kotlin `operator fun invoke`,
    /// Python `__call__`, PHP `__invoke`, Dart `call`, C# an `()` operator.
    ///
    /// All of them fill [`ProtocolSlot::Call`], so the call site probes ONE
    /// slot — this only says whether that probe is worth emitting. Declaring it
    /// replaces the `is_python_profile()` check that used to guard the probe,
    /// which is why `Counter()(3)` worked in Python and trapped with "Not a
    /// function" everywhere else.
    pub callable_objects: bool,

    /// `for x in obj` yields the object's KEYS for dict-like values (Map /
    /// Ordinary), while sequences (Array / Set / String) still yield values.
    /// Python `for k in dict` / JS-style object iteration. Routes the for-in
    /// values path through `collections::emit_iter_natural`.
    pub for_in_object_yields_keys: bool,

    /// Dict/hash literals (`{k: v}`) build a `Map` rather than an Ordinary
    /// object, so non-string keys keep their type (`{1: 'a'}` stays `1`, not
    /// `'1'`) and insertion order is preserved. Python dicts, Ruby hashes.
    pub dict_literals_as_map: bool,

    /// Stamp Python-style introspection metadata (`__name__`, `__mro__`) on each
    /// class object, so `Cls.__name__`, `Cls.__mro__`, and `type(obj).__name__`
    /// resolve. Off by default; languages with this reflection surface opt in.
    pub class_introspection_metadata: bool,

    /// Stamp `__fields` / `__methods` on each class object: the class's own
    /// members as reflection member tokens, the same 8-element shape
    /// `reflection::member_token_expr` produces.
    ///
    /// Without this there is NO runtime source for a class's member list —
    /// `stamp_reflection_type_fields` records only `__type`, `__typename` and
    /// `__kind`, and `__fields`/`__methods` exist solely inside an on-demand
    /// `ReflectionType` descriptor whose lists the caller supplies. That is why
    /// Pascal derives them into a compile-time HashMap and Dart's
    /// `reflection_adapter` builds its own: each language recomputes what
    /// nothing publishes.
    ///
    /// Off by default — it is real bytecode per class, so a language pays only
    /// when its reflection surface reads it.
    pub class_member_metadata: bool,


    /// Honour ALL declared bases (`NormalClass.bases`), computing a C3
    /// linearization and attaching every base's methods (multiple
    /// inheritance). Off by default: single-inheritance languages ignore
    /// `bases[1..]`, so their bytecode is unchanged. Python opts in.
    pub class_multiple_inheritance: bool,

    /// The `/` operator performs integer (truncating) division when BOTH
    /// operands are integers (C#, Java, C, Go). Languages where `/` is always
    /// real division (VB uses `\` for integer division; Python `/` is float)
    /// leave this false. Gates the integral-division lowering in
    /// `compile_binary`.
    pub integer_division_on_slash: bool,

    /// `xor` is ONE token with two meanings, resolved by operand type: bitwise
    /// when both operands are integers, logical otherwise (Pascal, and the same
    /// rule Delphi documents). The bitwise op is emitted either way; this only
    /// says whether a non-integer result is materialized as a Boolean.
    ///
    /// A property rather than a `(BuiltinType, ProtocolSlot)` binding, because
    /// it is not a question of WHICH emit target `xor` uses — both branches
    /// emit the same opcode — but of how the RESULT is typed. Recorded in
    /// builtinslotplan.md §3i: §3c mapped this site to "`BitXor` on `int`",
    /// which the code does not bear out.
    pub xor_is_logical_for_non_integers: bool,

    /// A `for` loop gives its loop variable a FRESH binding each iteration, so
    /// closures created in the body capture the per-iteration value (VB `For`,
    /// JS `let`-in-`for`). Languages where the loop variable is shared across
    /// iterations (C# / C `for`) leave this false.
    pub for_loop_per_iteration_binding: bool,

    /// A bare reference to a defined parameterless function INVOKES it (VB,
    /// Ruby) rather than yielding a function value. Languages where a bare name
    /// is a function reference leave this false.
    pub bare_name_invokes_parameterless_function: bool,

    /// Source functions are also published under an internal callable-value
    /// global so runtime string/name dispatch can resolve `"f"` to the same
    /// closure object as a direct `f(...)` call. Languages with PHP/Ruby-style
    /// dynamic callables opt in; the compiler code stays profile-driven.
    pub source_function_callable_aliases: bool,

    /// PHP: when a class constructor global is undefined at construction
    /// time, invoke the registered `spl_autoload_register` callback with
    /// the class name and retry. When false, a plain `GLOBAL_GET` is used.
    pub supports_autoload: bool,

    /// Language exposes buffered-iterator methods on generators
    /// (`current`/`next`/`valid`/`send`/`getReturn`/`throw`, PHP-style) and
    /// `foreach` must keep the generator's current value consistent with
    /// those methods. Drives the generic buffered-generator protocol (built
    /// on the WASM stack-switching `GEN_NEXT`/`RESUME` primitives). When
    /// false, generators use the plain advance-only iteration path.
    pub buffered_iterator_methods: bool,

    /// VB: LINQ query syntax compiled to method chains.
    pub linq_queries: bool,

    /// ECMA-262 §14.2 / §8.1: a block that declares a lexical binding
    /// (`let`/`const`/`class`) forms its own scope even when it contains
    /// *only* declarations — so `{ let x = 42; }` does not leak `x` to the
    /// enclosing scope. Languages without block-scoped lexical bindings (or
    /// whose blocks already scope unconditionally) leave this `false`.
    pub lexical_block_scope: bool,

    /// A variable binding belongs to the enclosing FUNCTION, not to the block
    /// it is written in: `if (true) { $x = 1; } echo $x;` is legal PHP.
    ///
    /// Deliberately NOT the same question as [`Self::lexical_block_scope`],
    /// which asks whether a block *containing* a `let`/`const` becomes its own
    /// scope. A language can block-scope without setting that one — C, Java,
    /// C#, Kotlin and Dart all leave it `false` — so it cannot stand in for
    /// this. Nor is it the same question as "does this language have a separate
    /// variable namespace" ([`vybe_runtime::registry::VariableNamespace`]);
    /// those two happen to coincide in PHP and are independent in general.
    ///
    /// Defaults `false` (block scoping), which is what every language got
    /// before this property existed.
    pub function_scoped_variables: bool,

    /// ECMA-262 §9.1.1.4.6 GetValue: reading an *unresolvable* reference (a
    /// name bound nowhere in the scope chain or on the global object) is a
    /// `ReferenceError` rather than yielding `undefined`. Statically-resolved
    /// languages and those that auto-vivify globals leave this `false`.
    pub unresolved_reference_throws: bool,

    /// Statically-typed languages coerce a value to its binding's declared
    /// type on assignment/initialization (C `_Bool b = 5` → `1`, int-width
    /// truncation, etc.). Dynamically-typed languages (JS, Python, PHP, …)
    /// infer a type *hint* for dispatch only and must never mutate the value,
    /// so they leave this `false`. Default `true` preserves the historical
    /// behaviour for the static languages.
    pub coerces_value_to_type_hint: bool,

    /// ECMA-262 §9.1.2 / §10.2.1.1: the receiver (`this`) is bound *ambiently*
    /// from the call context (Vybe carries it in the `__js_this` global the
    /// call site sets) rather than being passed as an explicit first positional
    /// parameter. When true, a method/constructor's arity excludes `this` and
    /// the body reads it from the context. Languages that thread the receiver
    /// as an explicit first parameter (`self`/`Me`/`$this` in slot 0) leave
    /// this `false`.
    pub ambient_this_binding: bool,

    /// namespaceplan.md migration switch: when true, this language's
    /// namespace-shaped entry points (host prefix chains, wildcard
    /// namespace member access, …) resolve through the common resolver
    /// (`primitives::resolver`) instead of the legacy hardcoded arms.
    /// Flipped per language as its phase lands (JS → Python → PHP →
    /// dotnet → rest); the legacy arms are deleted once every profile
    /// is migrated.
    pub uses_common_resolver: bool,

    /// ECMA-262 §10.2.1.1: a *missing* argument is the distinct value
    /// `undefined`, separate from an explicitly-passed `null`. Argument-presence
    /// tests (default-parameter application, constructor/overload arity
    /// dispatch) therefore key off "is undefined", so `f(null)` is recognized as
    /// one argument while `f()` is none. Languages with only a single
    /// nullish value leave this `false` and test "is null" instead.
    pub missing_arg_is_undefined: bool,

    /// When true, `StmtKind::ClassDecl` is routed through the new
    /// `common::classes::normalize_class` + `emit_class` path instead
    /// of the legacy `compile_class` orchestration. Enables the per-
    /// language migration from the classnormalization.md plan. Each
    /// language's walker opts in independently once its normalizer is
    /// implemented and tested. Default `false` → legacy path, zero
    /// behaviour change for unmigrated languages.
    pub uses_normalize_class: bool,

    /// Builtin function mappings: source name → emission action.
    pub builtins: HashMap<String, BuiltinDef>,

    /// Source type names this language uses for the built-in types, from the
    /// `[builtin_types]` section — `builtinslotplan.md` step 4.
    ///
    /// Consulted BEFORE the platform table in
    /// `vybe_ast::builtin_types::classify_with`, so a language both extends it
    /// (Python declaring `str`, which no shared list contains) and overrides it
    /// (a language where `real` means something other than a float).
    ///
    /// Empty for every language that has declared nothing, which is the
    /// neutral default: the platform table alone answers, exactly as before.
    ///
    /// This exists because step 3's census measured the platform's classifiers
    /// to be the binding constraint on the whole plan — a slot binding only
    /// ever applies where the receiver's type can be named, and PHP's `array`
    /// and Python's `str` reached those classifiers and failed them.
    pub builtin_type_spellings: Vec<vybe_ast::builtin_types::Spelling>,

    /// This language's OVERRIDES of the platform default slot table, from
    /// `[builtin_slots.<type>]` — `builtinslotplan.md` §2a.
    ///
    /// Empty for most languages, by design: §3a measured 32 of 36 string-slot
    /// cells already agreeing with ECMA. A language declares here only where it
    /// genuinely differs, and the platform default answers everything else.
    pub builtin_slots: vybe_ast::builtin_slots::BuiltinSlotBindings,

    /// Multi-opcode intrinsic definitions referenced by `emit = "intrinsic:<name>"`.
    pub intrinsics: HashMap<String, String>,

    /// Namespace resolution config.
    pub namespaces: NamespaceConfig,

    /// Known types: name → (host_module, host_function) for New TypeName().
    pub known_types: HashMap<String, (String, String)>,

    /// Value methods: instance methods called on values (str.toUpperCase(), arr.push()).
    /// The object is passed as first arg to the host function.
    /// Value methods can have multiple overloads by arity.
    /// E.g. `Add(item)` for list (1 arg) vs `Add(key, value)` for dict (2 args).
    pub value_methods: HashMap<String, Vec<BuiltinDef>>,

    /// Namespace constants: property access that returns a value, NOT a function call.
    /// "Math.PI" → 3.14159..., "Number.MAX_SAFE_INTEGER" → 9007199254740991
    pub namespace_constants: HashMap<String, ConstantValue>,

    /// Array higher-order methods routed to compiled JS builtins.
    /// "map" → "__array_map", "filter" → "__array_filter", etc.
    pub array_methods: HashMap<String, String>,

    /// Return types of builtin free functions, for type inference on their
    /// result (e.g. VB `Command`/`Environ` → "String", `Timer` → "Double"),
    /// keyed by lowercased name. Data-driven so no language-name check is
    /// needed in `infer_function_return_type`.
    pub builtin_return_types: HashMap<String, String>,

    /// Free functions that extract a field from a `DateTime` argument
    /// (VB `Year(d)` → `d.Year`, `Month(d)` → `d.Month`, …), keyed by lowercased
    /// function name → the DateTime field. Applied only when the argument is a
    /// DateTime (compile-time type check stays in the compiler); the profile
    /// supplies WHICH names, so no language-name check is needed.
    pub datetime_field_functions: HashMap<String, String>,

    /// Synthetic ESM imports the language treats as pre-declared (i.e.
    /// ambient) at module scope. The Linker walks these BEFORE user
    /// imports so `import { X }` in user code shadows a profile default
    /// with the same local name (ECMA-262 §16.2 lexical-over-module-scope
    /// rule).
    ///
    /// Declared in profile TOML as `[[esm_default]]` entries in one
    /// of three shapes: `kind = "named"`, `kind = "namespace"`, or
    /// `kind = "package-root"`.
    pub esm_defaults: Vec<EsmDefault>,

    /// Bare-import-specifier rewrites for languages with built-in
    /// modules that resolve without a prefix. JS/Node: `import fs
    /// from 'fs'` rewrites to `'node:fs'` so it binds the same host
    /// module as the explicit `'node:fs'` import.
    ///
    /// Languages with their own stdlib semantics for these names
    /// (Python `import os`, Ruby `require 'fileutils'`, …) leave the
    /// table empty so their bare imports stay unrouted by this
    /// mechanism. Declared in profile TOML as a `[bare_module_aliases]`
    /// table: `"fs" = "node:fs"` etc.
    pub bare_module_aliases: HashMap<String, String>,
}

/// One pre-declared ESM import in the ambient module scope — the
/// profile's equivalent of a hand-written `import` statement. Three
/// variants mirror the ECMA-262 import forms.
#[derive(Debug, Clone)]
pub enum EsmDefault {
    /// `import { name as local } from "module"`. `name` defaults to
    /// `local` when not provided.
    Named {
        local: String,
        module: String,
        name: String,
    },
    /// `import * as alias from "module"`. Qualified access `alias.field`
    /// resolves to `(module, field)` at compile time.
    Namespace { alias: String, module: String },
    /// Component-Model package root. A qualified chain whose first
    /// segment matches `prefix` maps to a specifier built by joining
    /// `module_root` + remaining-but-last segments + `/` + last segment.
    /// Used for idiomatic qualified access (VB `Imports System` →
    /// `System.Threading.Thread.Sleep` resolves under `dotnet:`) where
    /// the namespace object would be too coarse.
    PackageRoot { prefix: String, module_root: String },
    /// Mount-with-rename export surface for a language-level module
    /// (namespaceplan.md alias leaves as profile data): declares that
    /// module `module` exports `name`, implemented by
    /// `target_module`/`target_name`. Feeds the Linker's
    /// `module_exports` re-export map, so `from json import dumps`
    /// (Python) binds `dumps` → `ecma:json`/`stringify` through the
    /// SAME Named-import path every ESM import takes — reconciling
    /// source names with the canonical host export names.
    ModuleExport {
        module: String,
        name: String,
        target_module: String,
        target_name: String,
    },
    /// Namespace-tree mount (namespaceplan.md): a qualified chain whose
    /// first segment matches `prefix` resolves by walking the global
    /// namespace tree rooted at `path` — `System.Math.Sin` with
    /// `prefix = "system", path = "dotnet.system"` walks
    /// `dotnet.system.math.sin`. This is how a language's ambient
    /// namespace roots map onto platform-registered tree surfaces
    /// (VB/C# `System.*` → the dotnet descriptor data) with zero
    /// platform-owned resolver logic.
    TreeMount { prefix: String, path: String },
    /// Ambient namespace-tree root (namespaceplan.md): bare qualified
    /// chains additionally resolve by searching under `path` — the
    /// data form of .NET's ambient `Imports`/`using` context
    /// (`Thread.Sleep` found at `dotnet.system.threading.thread.sleep`).
    /// Profile entries give the language's defaults; user `Imports X.Y`
    /// statements add more at link time (rebased through the tree-mounts).
    TreeAmbient { path: String },
}

/// A compile-time constant value.
#[derive(Debug, Clone)]
pub enum ConstantValue {
    Float(f64),
    Str(String),
}

/// String indexing style.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum StringIndexing {
    ZeroBased,
    OneBased,
}

/// Namespace resolution configuration.
#[derive(Debug, Clone, Default)]
pub struct NamespaceConfig {
    /// Enable .NET runtime surface semantics that are not namespace lookup
    /// decisions (collection dispatch, primitive aliases, WinForms inference).
    /// Namespace lookup itself is driven by `uses_namespace_resolver()` and
    /// the tree/default import data below.
    pub use_dotnet: bool,
    /// Treat simple source imports (`Imports Demo.Core`, `using Demo.Core`)
    /// as namespace imports for user declarations too. This is generic source
    /// namespace behavior, separate from platform tree mounts like
    /// `System -> dotnet.system`.
    pub source_imports_are_namespaces: bool,
    /// Additional imports beyond the defaults (e.g. "microsoft.visualbasic" for VB).
    pub extra_imports: Vec<String>,
    /// Known namespace roots (used when use_dotnet is false).
    pub roots: Vec<String>,
    /// Default imports always available (used when use_dotnet is false).
    pub default_imports: Vec<String>,
    /// Known constants (property access, not function call).
    pub constants: Vec<String>,
    /// Namespace-tree ROOTS whose registered types this language resolves
    /// member access against (`["dotnet"]`, `["flutter"]`, `["gcl"]`).
    ///
    /// A statically-typed receiver (`Button b; b.Text`) resolves its members
    /// through the platform's `Type` node — its properties, its methods, its
    /// constructor. The same roots are mounted AMBIENTLY, so an unqualified
    /// name (`Scaffold(...)`, `TForm`, `new Button()`) resolves to its `Type`
    /// and constructs through the one common-resolver `Ctor` path. That is why
    /// a GUI platform needs no compiler-side registration pass: declaring the
    /// root is the whole integration. Naming the roots is what makes that language-neutral: the
    /// question is never "is this C#", it is "which catalogs did this language
    /// import". Empty means the language has no platform types.
    pub type_scopes: Vec<String>,
    /// Namespace-tree path under which registered types are dispatched at
    /// RUNTIME through the type registry rather than through the language's
    /// value-method tables (`["dotnet", "system", "collections"]`).
    ///
    /// A PATH, not a language name: the question "should `xs.Add(v)` defer to
    /// runtime dispatch" is answered by where the declaring type lives, so a
    /// language opts in by naming the subtree. Empty means never defer.
    pub runtime_collection_scope: Vec<String>,
}

/// How a function returns its value.
#[derive(Debug, Clone, PartialEq)]
pub enum ReturnStyle {
    /// Assign to a Result slot; function epilogue returns it. (Pascal)
    ResultSlot,
    /// Explicit `return expr` statements. (Python, JS, C#, PHP, VB)
    Explicit,
    /// Last expression in body is the return value. (Ruby)
    LastExpression,
}

/// How a language forms reflection / runtime type-identity names. Owned by
/// the profile so each language keeps its own type namespace — see
/// [`Profile::reflection_type_naming`]. Extend with a new variant when a
/// language needs a distinct scheme (e.g. a `java.lang.*` root).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ReflectionTypeNaming {
    /// A type keeps its own name (`Throwable`, `Error`, `MyClass`). No forced
    /// namespace, no BCL primitive remap. Default for every language except
    /// the .NET family — this is what preserves each language's real
    /// exception hierarchy in its `__types` chain.
    #[default]
    Native,
    /// .NET BCL scheme: qualify under `System.` and map primitives
    /// (`int`→`Int32`, `string`→`String`, …). C#/VB reflection expects this.
    Dotnet,
}

/// Definition of a builtin function's compilation.
#[derive(Debug, Clone)]
pub struct BuiltinDef {
    pub emit: BuiltinEmit,
    pub min_args: u8,
    pub max_args: u8,

    /// The protocol slot this method implements, from `slot = "len"` —
    /// `builtinslotplan.md` step 4b.
    ///
    /// This is how the platform learns that Dart's `length`, Python's `__len__`
    /// and PHP's `count` are all the same operation, WITHOUT a method-name
    /// table in shared code: the language owns its spellings and declares which
    /// shared slot each one fills. Step 3's census could only record method
    /// names for exactly this reason — mapping name → slot centrally is the
    /// anti-pattern the plan exists to remove.
    ///
    /// `None` — the overwhelming default — means the method keeps its declared
    /// `emit` and no slot resolution is attempted.
    pub slot: Option<vybe_ast::ProtocolSlot>,
}

/// What to emit for a builtin call.
#[derive(Debug, Clone)]
pub enum BuiltinEmit {
    /// Call a host import: (module, name)
    HostCall(String, String),
    /// Emit a direct opcode sequence by name.
    Opcode(String),
    /// Mutate a variable: var = var OP arg. (Inc, Dec)
    MutateVar(String), // "add" or "sub"
    /// Multi-opcode intrinsic: name references [intrinsics] table in profile.
    Intrinsic(String),
    /// Dispatch to a compiler_common opcode-style emitter (args already on stack).
    /// e.g. "dict.set_dynamic", "collections.push", "strings.length"
    Common(String),
    /// Print (variadic)
    Print,
    /// String length
    StrLength,
    /// Emit nothing (no-op, e.g. randomize, free)
    Noop,
    /// Dynamic method dispatch via `ecma:value.invokeMethod`.
    /// Used for methods with polymorphic receiver types (JS `str.slice` vs
    /// `arr.slice`) — runtime picks the right implementation based on the
    /// receiver. Args are compiled as `[receiver, arg1, ..., argN]` and
    /// the emitter splices in the method name.
    Invoke(String),
}

impl LanguageProfile {
    /// Whether this profile uses the shared namespace resolver.
    ///
    /// This is data-shaped, not platform-shaped: a language opts in through
    /// common resolver imports, namespace-tree mounts/ambients, type scopes,
    /// package roots, or source namespace roots. `dotnet` is only one such
    /// mounted tree, same as `flutter`, `plib`, `php`, etc.
    pub fn uses_namespace_resolver(&self) -> bool {
        self.uses_common_resolver
            || !self.namespaces.type_scopes.is_empty()
            || !self.esm_defaults.is_empty()
            || !self.namespaces.roots.is_empty()
            || !self.namespaces.default_imports.is_empty()
    }

    /// Look up a builtin by name (case-insensitive for case-insensitive languages).
    pub fn lookup_builtin(&self, name: &str) -> Option<&BuiltinDef> {
        if self.case_sensitive {
            self.builtins.get(name).or_else(|| {
                if self.name == "php" {
                    let lower = name.to_lowercase();
                    self.builtins.get(&lower)
                } else {
                    None
                }
            })
        } else {
            let lower = name.to_lowercase();
            self.builtins.get(&lower)
        }
    }

    /// Look up a known type constructor mapping.
    pub fn lookup_known_type(&self, name: &str) -> Option<(&str, &str)> {
        let key = if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        };
        self.known_types
            .get(&key)
            .map(|(m, f)| (m.as_str(), f.as_str()))
    }

    /// Check if a name is a known namespace root.
    pub fn is_namespace_root(&self, name: &str) -> bool {
        let key = if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        };
        self.namespaces.roots.iter().any(|r| r == &key)
    }

    /// Check if a name is a known constant (property access, not call).
    pub fn is_namespace_constant(&self, name: &str) -> bool {
        let key = if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        };
        self.namespaces.constants.iter().any(|c| c == &key)
    }

    /// Look up a value method by name + arity.
    /// Returns the first overload whose arity range matches.
    pub fn lookup_value_method(&self, name: &str, argc: u8) -> Option<&BuiltinDef> {
        let key = if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        };
        let overloads = self.value_methods.get(&key)?;
        overloads
            .iter()
            .find(|d| argc >= d.min_args && argc <= d.max_args)
    }

    /// Check if a value method exists by name (any arity).
    pub fn has_value_method(&self, name: &str) -> bool {
        let key = if self.case_sensitive {
            name.to_string()
        } else {
            name.to_lowercase()
        };
        self.value_methods.contains_key(&key)
    }

    /// Look up a namespace constant value (Math.PI, Number.MAX_SAFE_INTEGER).
    pub fn lookup_constant(&self, name: &str) -> Option<&ConstantValue> {
        if self.case_sensitive {
            self.namespace_constants.get(name)
        } else {
            self.namespace_constants.get(&name.to_lowercase())
        }
    }

    /// Look up an array method routing (map → __array_map).
    pub fn lookup_array_method(&self, name: &str) -> Option<&str> {
        self.array_methods.get(name).map(|s| s.as_str())
    }
}

/// Parse a TOML profile source into a LanguageProfile.
/// Parse a language profile, caching the result. Profiles are constant
/// `&'static str` (via `include_str!`), but every test re-parses the ~900-line
/// TOML — a per-test cost EVERY language pays (Python/C# have no prelude to
/// cache, so this is their main lever). Keyed by the source's (ptr, len),
/// which is stable and unique for `&'static str`; cloning a `LanguageProfile`
/// is far cheaper than re-parsing TOML + rebuilding its tables.
/// Platform-registered namespace constants for `use_dotnet` profiles
/// (name → value). The dotnet platform seeds these at startup (and, once it is
/// a loadable module, at load time) so this generic parser never references the
/// platform — the same inversion every future platform/language dylib uses.
static DOTNET_NS_CONSTANTS: std::sync::OnceLock<Vec<(String, f64)>> = std::sync::OnceLock::new();

/// Called once by the host/compiler with `dotnet::namespace_constant_mappings()`.
pub fn register_dotnet_namespace_constants(mappings: Vec<(String, f64)>) {
    let _ = DOTNET_NS_CONSTANTS.set(mappings);
}

fn dotnet_namespace_constants() -> &'static [(String, f64)] {
    DOTNET_NS_CONSTANTS.get().map(Vec::as_slice).unwrap_or(&[])
}

/// Turn an emit-target string into a [`BuiltinEmit`].
///
/// The profile's emit-target vocabulary — `opcode:` / `intrinsic:` / `common:` /
/// `host:` / `invoke:` / `mutate:` / `print` / `noop` — in one place.
///
/// Public because `builtinslotplan.md` step 4b resolves a slot binding to one
/// of these same strings and needs to turn it into an emit. Sharing this parser
/// is the point: a slot binding is deliberately NOT a new vocabulary, so it must
/// not get a second, drifting interpreter.
pub fn parse_emit_target(s: &str) -> Option<BuiltinEmit> {
    match s {
        "print" => Some(BuiltinEmit::Print),
        "str_length" => Some(BuiltinEmit::StrLength),
        "noop" => Some(BuiltinEmit::Noop),
        _ if s.starts_with("host:") => {
            let parts: Vec<&str> = s["host:".len()..].splitn(3, ':').collect();
            if parts.len() == 3 {
                Some(BuiltinEmit::HostCall(
                    format!("{}:{}", parts[0], parts[1]),
                    parts[2].to_string(),
                ))
            } else {
                None
            }
        }
        _ if s.starts_with("opcode:") => {
            Some(BuiltinEmit::Opcode(s["opcode:".len()..].to_string()))
        }
        _ if s.starts_with("mutate:") => {
            Some(BuiltinEmit::MutateVar(s["mutate:".len()..].to_string()))
        }
        _ if s.starts_with("intrinsic:") => {
            Some(BuiltinEmit::Intrinsic(s["intrinsic:".len()..].to_string()))
        }
        _ if s.starts_with("common:") => {
            Some(BuiltinEmit::Common(s["common:".len()..].to_string()))
        }
        _ if s.starts_with("invoke:") => {
            Some(BuiltinEmit::Invoke(s["invoke:".len()..].to_string()))
        }
        _ => None,
    }
}

pub fn parse_profile(src: &str) -> Result<LanguageProfile, String> {
    use std::collections::HashMap;
    use std::sync::{Mutex, OnceLock};
    // Keyed by CONTENT (not pointer): correct even if a caller passes a
    // dynamically-built profile string. Hashing ~900 chars is still far
    // cheaper than re-parsing the TOML and rebuilding the profile's tables.
    static CACHE: OnceLock<Mutex<HashMap<String, LanguageProfile>>> = OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    if let Some(p) = cache.lock().unwrap().get(src) {
        return Ok(p.clone());
    }
    let parsed = parse_profile_uncached(src)?;
    cache
        .lock()
        .unwrap()
        .insert(src.to_string(), parsed.clone());
    Ok(parsed)
}

fn parse_profile_uncached(src: &str) -> Result<LanguageProfile, String> {
    use toml::Value;

    let root: Value =
        toml::from_str(src).map_err(|e| format!("TOML parse error in profile: {}", e))?;

    let compiler = root.get("compiler").ok_or("Missing [compiler] section")?;

    let name = root
        .get("info")
        .and_then(|v| v.get("name"))
        .and_then(|v| v.as_str())
        .unwrap_or("")
        .to_string();

    let function_return = match compiler
        .get("function_return")
        .and_then(|v| v.as_str())
        .unwrap_or("explicit")
    {
        "result_slot" => ReturnStyle::ResultSlot,
        "last_expression" => ReturnStyle::LastExpression,
        _ => ReturnStyle::Explicit,
    };

    let result_slot_name = compiler
        .get("result_slot_name")
        .and_then(|v| v.as_str())
        .unwrap_or("Result")
        .to_string();

    let self_keyword = compiler
        .get("self_keyword")
        .and_then(|v| v.as_str())
        .unwrap_or("this")
        .to_string();
    let base_keyword = compiler
        .get("base_keyword")
        .and_then(|v| v.as_str())
        .map(|s| s.to_string());
    let constructor_name = compiler
        .get("constructor_name")
        .and_then(|v| v.as_str())
        .unwrap_or("constructor")
        .to_string();
    let member_call_on_null_error = compiler
        .get("member_call_on_null_error")
        .and_then(|v| v.as_str())
        .unwrap_or("TypeError")
        .to_string();
    let class_method_dispatch = compiler
        .get("class_method_dispatch")
        .and_then(|v| v.as_str())
        .unwrap_or("instance")
        .to_string();
    let dynamic_numeric_dispatch = compiler
        .get("dynamic_numeric_dispatch")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let enum_as_ordinals = compiler
        .get("enum_as_ordinals")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let case_sensitive = compiler
        .get("case_sensitive")
        .and_then(|v| v.as_bool())
        .unwrap_or(true);
    let string_indexing = match compiler
        .get("string_indexing")
        .and_then(|v| v.as_str())
        .unwrap_or("zero_based")
    {
        "one_based" => StringIndexing::OneBased,
        _ => StringIndexing::ZeroBased,
    };
    let array_upper_bound_inclusive = compiler
        .get("array_upper_bound_inclusive")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let negative_index_wraps = compiler
        .get("negative_index_wraps")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let slice_step_zero_raises = compiler
        .get("slice_step_zero_raises")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let tuple_literals_tagged = compiler
        .get("tuple_literals_tagged")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let parens_for_index = compiler
        .get("parens_for_index")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let unified_array_map = compiler
        .get("unified_array_map")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let concat_stringifies_operands = compiler
        .get("concat_stringifies_operands")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let entry_point = compiler
        .get("entry_point")
        .and_then(|v| v.as_str())
        .map(|s| s.to_string());
    let hoist_var = compiler
        .get("hoist_var")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let dynamic_add = compiler
        .get("dynamic_add")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let function_references = compiler
        .get("function_references")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let commonjs_require = compiler
        .get("commonjs_require")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let multi_value_tuple_returns = compiler
        .get("multi_value_tuple_returns")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let methods_bind_on_access = compiler
        .get("methods_bind_on_access")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let default_args_evaluated_once = compiler
        .get("default_args_evaluated_once")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let byref_boxing = compiler
        .get("byref_boxing")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let with_block = compiler
        .get("with_block")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let new_with_initializer = compiler
        .get("new_with_initializer")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let new_from_initializer = compiler
        .get("new_from_initializer")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let linq_queries = compiler
        .get("linq_queries")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let switch_fallthrough = compiler
        .get("switch_fallthrough")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let throwable_is_root = compiler
        .get("throwable_is_root")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let field_hiding = compiler
        .get("field_hiding")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let methods_virtual_by_default = compiler
        .get("methods_virtual_by_default")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let separate_property_method_namespace = compiler
        .get("separate_property_method_namespace")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let reflection_type_naming = match compiler
        .get("reflection_type_naming")
        .and_then(|v| v.as_str())
        .map(|s| s.to_ascii_lowercase())
        .as_deref()
    {
        Some("dotnet") => ReflectionTypeNaming::Dotnet,
        // Default (and explicit "native"): each language owns its type names.
        _ => ReflectionTypeNaming::Native,
    };
    let supports_private_fields = compiler
        .get("supports_private_fields")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_error_object_shape = compiler
        .get("ecma_error_object_shape")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_undefined_value = compiler
        .get("has_undefined_value")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let supports_spread_arguments = compiler
        .get("supports_spread_arguments")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let supports_dynamic_import = compiler
        .get("supports_dynamic_import")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_ecma_bigint = compiler
        .get("has_ecma_bigint")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_strict_mode = compiler
        .get("ecma_strict_mode")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_switch_strict_equality = compiler
        .get("ecma_switch_strict_equality")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_lexical_declarations = compiler
        .get("ecma_lexical_declarations")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_ecma_globals = compiler
        .get("has_ecma_globals")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_operator_coercion = compiler
        .get("ecma_operator_coercion")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let abstract_equality = compiler
        .get("abstract_equality")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_arguments_object = compiler
        .get("has_arguments_object")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let relaxed_call_arity = compiler
        .get("relaxed_call_arity")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_new_dispatch = compiler
        .get("ecma_new_dispatch")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_array_elisions = compiler
        .get("ecma_array_elisions")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_generators = compiler
        .get("has_generators")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_object_literals = compiler
        .get("ecma_object_literals")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_boolean_operators = compiler
        .get("ecma_boolean_operators")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let dynamic_member_access = compiler
        .get("dynamic_member_access")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_typeof_operator = compiler
        .get("ecma_typeof_operator")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_to_primitive = compiler
        .get("ecma_to_primitive")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_in_operator = compiler
        .get("ecma_in_operator")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_array_method_dispatch = compiler
        .get("ecma_array_method_dispatch")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_promise_methods = compiler
        .get("ecma_promise_methods")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let ecma_iterator_result_shape = compiler
        .get("ecma_iterator_result_shape")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let async_wraps_body_in_try = compiler
        .get("async_wraps_body_in_try")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_function_constructor = compiler
        .get("has_function_constructor")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let has_function_prototype_bind = compiler
        .get("has_function_prototype_bind")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let function_invocation_members = compiler
        .get("function_invocation_members")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let string_aware_relational = compiler
        .get("string_aware_relational")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let function_scoped_variables = compiler
        .get("function_scoped_variables")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let lexical_block_scope = compiler
        .get("lexical_block_scope")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let unresolved_reference_throws = compiler
        .get("unresolved_reference_throws")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let coerces_value_to_type_hint = compiler
        .get("coerces_value_to_type_hint")
        .and_then(|v| v.as_bool())
        .unwrap_or(true);
    let ambient_this_binding = compiler
        .get("ambient_this_binding")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let uses_common_resolver = compiler
        .get("uses_common_resolver")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let missing_arg_is_undefined = compiler
        .get("missing_arg_is_undefined")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let materialize_bool_results = compiler
        .get("materialize_bool_results")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let callable_objects = compiler
        .get("callable_objects")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let for_in_object_yields_keys = compiler
        .get("for_in_object_yields_keys")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let dict_literals_as_map = compiler
        .get("dict_literals_as_map")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let class_introspection_metadata = compiler
        .get("class_introspection_metadata")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let class_member_metadata = compiler
        .get("class_member_metadata")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let class_multiple_inheritance = compiler
        .get("class_multiple_inheritance")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let integer_division_on_slash = compiler
        .get("integer_division_on_slash")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let xor_is_logical_for_non_integers = compiler
        .get("xor_is_logical_for_non_integers")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let for_loop_per_iteration_binding = compiler
        .get("for_loop_per_iteration_binding")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let bare_name_invokes_parameterless_function = compiler
        .get("bare_name_invokes_parameterless_function")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let source_function_callable_aliases = compiler
        .get("source_function_callable_aliases")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let supports_autoload = compiler
        .get("supports_autoload")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let buffered_iterator_methods = compiler
        .get("buffered_iterator_methods")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let uses_normalize_class = compiler
        .get("uses_normalize_class")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);

    fn parse_builtin_table(root: &Value, section: &str) -> HashMap<String, BuiltinDef> {
        let mut map = HashMap::new();
        if let Some(bt) = root.get(section).and_then(|v| v.as_table()) {
            for (name, val) in bt {
                if let Some(t) = val.as_table() {
                    let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
                    let min_args =
                        t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
                    let max_args = t
                        .get("max_args")
                        .and_then(|v| v.as_integer())
                        .unwrap_or(255) as u8;
                    let slot = parse_slot(t);
                    if let Some(emit) = parse_emit(emit_str) {
                        map.insert(
                            name.clone(),
                            BuiltinDef {
                                emit,
                                min_args,
                                max_args,
                                slot,
                            },
                        );
                    }
                }
            }
        }
        map
    }

    /// Read a `slot = "len"` declaration off a builtin/value-method entry —
    /// `builtinslotplan.md` step 4b.
    ///
    /// An unrecognised name yields `None`, deliberately: a profile written
    /// against a newer toolchain must still load, with the unknown declaration
    /// simply having no effect. Failing the whole profile over one unknown slot
    /// would make adding a slot a breaking change for every language.
    fn parse_slot(t: &toml::value::Table) -> Option<vybe_ast::ProtocolSlot> {
        t.get("slot")
            .and_then(|v| v.as_str())
            .and_then(vybe_ast::ProtocolSlot::from_key)
    }

    fn parse_emit(s: &str) -> Option<BuiltinEmit> {
        parse_emit_target(s)
    }

    /// Parse `[builtin_types]` — `builtinslotplan.md` step 4.
    ///
    /// ```toml
    /// [builtin_types]
    /// string = ["str"]                 # Python
    /// array  = ["array", "list*"]      # PHP / a language with List<T>
    /// map    = ["dict", "*dictionary*"]
    /// ```
    ///
    /// The key is a `BuiltinType` key (`string`, `int`, `array`, …). Each entry
    /// is a spelling, with `*` marking where it need not match:
    ///
    /// | written | matches |
    /// |---|---|
    /// | `str` | the whole hint equals `str` |
    /// | `*.string` | the hint ENDS with `.string` |
    /// | `list<*` | the hint STARTS with `list<` |
    /// | `*dictionary*` | the hint CONTAINS `dictionary` |
    ///
    /// `*` is deliberately not a glob: three fixed shapes, matching the three
    /// the platform table already uses, so a profile cannot express a matcher
    /// the shared table has no way to run.
    ///
    /// An unknown type key or a bare `*` is skipped rather than failing the
    /// profile — a profile that names a built-in this version does not have
    /// should still load, and the type simply stays unresolvable.
    fn parse_builtin_types(root: &Value) -> Vec<vybe_ast::builtin_types::Spelling> {
        use vybe_ast::builtin_slots::BuiltinType;
        use vybe_ast::builtin_types::{Match, Spelling};

        let mut out = Vec::new();
        let Some(table) = root.get("builtin_types").and_then(|v| v.as_table()) else {
            return out;
        };
        for (type_key, entries) in table {
            let Some(ty) = BuiltinType::from_key(type_key) else {
                continue;
            };
            let Some(list) = entries.as_array() else {
                continue;
            };
            for entry in list {
                let Some(raw) = entry.as_str() else { continue };
                let leading = raw.starts_with('*');
                let trailing = raw.ends_with('*') && raw.len() > 1;
                let core = raw.trim_start_matches('*').trim_end_matches('*');
                if core.is_empty() {
                    continue;
                }
                let how = match (leading, trailing) {
                    (true, true) => Match::Contains,
                    (true, false) => Match::Suffix,
                    (false, true) => Match::Prefix,
                    (false, false) => Match::Exact,
                };
                out.push(Spelling::owned(core, how, ty));
            }
        }
        out
    }

    /// `[builtin_slots.<type>]` — the language's OVERRIDES of the platform
    /// default `(BuiltinType, ProtocolSlot) -> emit target` table.
    ///
    /// ```toml
    /// [builtin_slots.map]
    /// get_item = "common:dart.index_get"   # Dart: a miss is `null`, not `undefined`
    ///
    /// [builtin_slots.string]
    /// len = "common:str_length"            # PHP counts bytes
    /// ```
    ///
    /// This is `builtinslotplan.md` §2a's shape, and it is what makes a slot
    /// bindable when languages genuinely DISAGREE. Measured 2026-07-31: the
    /// central table could not bind `Map`/`GetItem` because a missing key is
    /// `undefined` in JS, `null` in Dart, a `KeyError` in Python and
    /// `null`-plus-a-warning in PHP — four right answers, so any single central
    /// entry encodes one language's as everyone's. Same for `Eq`, where
    /// Python's set equality is order-independent and Dart's record equality is
    /// structural.
    ///
    /// Precedence is the language's table first, then the platform default —
    /// §2d steps 2 and 3, implemented by `BuiltinSlotBindings::get_or`.
    ///
    /// # An unrecognised key here is FATAL, unlike `[builtin_types]`
    ///
    /// `parse_builtin_types` skips what it does not recognise, and the cost of
    /// that is bounded: a dropped spelling leaves the type unresolvable and the
    /// language's own emitter keeps running.
    ///
    /// A dropped *override* is not bounded. It silently falls through to the
    /// platform default, which is precisely the `undefined`-instead-of-`null`
    /// failure that forced `Map`/`GetItem` to be backed out on 2026-07-31 — and
    /// Dart's correctness now depends on exactly one entry in this table. A
    /// one-character typo (`get_itm`, `[builtin_slots.mop]`) would reintroduce
    /// that bug with no signal anywhere.
    ///
    /// Profiles ship WITH the platform in this repo, so "written against a
    /// newer version" is not a real case — an unrecognised key is a typo, every
    /// time. It fails the profile loudly.
    fn parse_builtin_slots(
        root: &Value,
    ) -> Result<vybe_ast::builtin_slots::BuiltinSlotBindings, String> {
        use vybe_ast::builtin_slots::{BuiltinSlotBindings, BuiltinType};

        let mut out = BuiltinSlotBindings::new();
        let Some(table) = root.get("builtin_slots").and_then(|v| v.as_table()) else {
            return Ok(out);
        };
        for (type_key, entries) in table {
            let ty = BuiltinType::from_key(type_key).ok_or_else(|| {
                format!("[builtin_slots.{type_key}]: `{type_key}` is not a built-in type")
            })?;
            let slots = entries.as_table().ok_or_else(|| {
                format!("[builtin_slots.{type_key}] must be a table of slot = \"target\"")
            })?;
            for (slot_key, target) in slots {
                let slot = vybe_ast::ProtocolSlot::from_key(slot_key).ok_or_else(|| {
                    format!("[builtin_slots.{type_key}]: `{slot_key}` is not a protocol slot")
                })?;
                let target = target.as_str().ok_or_else(|| {
                    format!("[builtin_slots.{type_key}] {slot_key} must be an emit-target string")
                })?;
                out.insert(ty, slot, target);
            }
        }
        Ok(out)
    }

    fn parse_string_table(root: &Value, section: &str) -> HashMap<String, String> {
        let mut map = HashMap::new();
        if let Some(t) = root.get(section).and_then(|v| v.as_table()) {
            for (k, v) in t {
                if let Some(s) = v.as_str() {
                    map.insert(k.clone(), s.to_string());
                }
            }
        }
        map
    }

    fn parse_value_methods_table(root: &Value) -> HashMap<String, Vec<BuiltinDef>> {
        let mut map: HashMap<String, Vec<BuiltinDef>> = HashMap::new();
        if let Some(bt) = root.get("value_methods").and_then(|v| v.as_table()) {
            for (name, val) in bt {
                // Either a single inline table or an array of inline tables (overloads)
                if let Some(arr) = val.as_array() {
                    for entry in arr {
                        if let Some(t) = entry.as_table() {
                            let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
                            let min_args =
                                t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
                            let max_args = t
                                .get("max_args")
                                .and_then(|v| v.as_integer())
                                .unwrap_or(255) as u8;
                            let slot = parse_slot(t);
                            if let Some(emit) = parse_emit(emit_str) {
                                map.entry(name.clone()).or_default().push(BuiltinDef {
                                    emit,
                                    min_args,
                                    max_args,
                                    slot,
                                });
                            }
                        }
                    }
                } else if let Some(t) = val.as_table() {
                    let emit_str = t.get("emit").and_then(|v| v.as_str()).unwrap_or("noop");
                    let min_args =
                        t.get("min_args").and_then(|v| v.as_integer()).unwrap_or(0) as u8;
                    let max_args = t
                        .get("max_args")
                        .and_then(|v| v.as_integer())
                        .unwrap_or(255) as u8;
                    let slot = parse_slot(t);
                    if let Some(emit) = parse_emit(emit_str) {
                        map.entry(name.clone()).or_default().push(BuiltinDef {
                            emit,
                            min_args,
                            max_args,
                            slot,
                        });
                    }
                }
            }
        }
        map
    }

    let mut builtins = parse_builtin_table(&root, "builtins");
    let mut value_methods = parse_value_methods_table(&root);
    let intrinsics = parse_string_table(&root, "intrinsics");
    let builtin_return_types: HashMap<String, String> =
        parse_string_table(&root, "builtin_return_types")
            .into_iter()
            .map(|(k, v)| (k.to_lowercase(), v))
            .collect();
    let datetime_field_functions: HashMap<String, String> =
        parse_string_table(&root, "datetime_field_functions")
            .into_iter()
            .map(|(k, v)| (k.to_lowercase(), v))
            .collect();
    let mut array_methods = parse_string_table(&root, "array_methods");

    let namespaces = if let Some(ns) = root.get("namespaces") {
        NamespaceConfig {
            use_dotnet: ns
                .get("use_dotnet")
                .and_then(|v| v.as_bool())
                .unwrap_or(false),
            source_imports_are_namespaces: ns
                .get("source_imports_are_namespaces")
                .and_then(|v| v.as_bool())
                .unwrap_or(false),
            extra_imports: ns
                .get("extra_imports")
                .and_then(|v| v.as_array())
                .map(|a| {
                    a.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect()
                })
                .unwrap_or_default(),
            roots: ns
                .get("roots")
                .and_then(|v| v.as_array())
                .map(|a| {
                    a.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect()
                })
                .unwrap_or_default(),
            default_imports: ns
                .get("default_imports")
                .and_then(|v| v.as_array())
                .map(|a| {
                    a.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect()
                })
                .unwrap_or_default(),
            type_scopes: ns
                .get("type_scopes")
                .and_then(|v| v.as_array())
                .map(|a| {
                    a.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect()
                })
                .unwrap_or_default(),
            runtime_collection_scope: ns
                .get("runtime_collection_scope")
                .and_then(|v| v.as_array())
                .map(|a| {
                    a.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect()
                })
                .unwrap_or_default(),
            constants: ns
                .get("constants")
                .and_then(|v| v.as_array())
                .map(|a| {
                    a.iter()
                        .filter_map(|v| v.as_str().map(|s| s.to_string()))
                        .collect()
                })
                .unwrap_or_default(),
        }
    } else {
        NamespaceConfig::default()
    };

    let mut known_types = HashMap::new();
    if let Some(kt) = root.get("known_types").and_then(|v| v.as_table()) {
        for (name, val) in kt {
            if let Some(arr) = val.as_array() {
                if arr.len() == 2 {
                    if let (Some(m), Some(f)) = (arr[0].as_str(), arr[1].as_str()) {
                        known_types.insert(name.clone(), (m.to_string(), f.to_string()));
                    }
                }
            }
        }
    }

    if !case_sensitive {
        builtins = builtins
            .into_iter()
            .map(|(name, def)| (name.to_lowercase(), def))
            .collect();
        value_methods = value_methods
            .into_iter()
            .map(|(name, defs)| (name.to_lowercase(), defs))
            .collect();
        array_methods = array_methods
            .into_iter()
            .map(|(name, target)| (name.to_lowercase(), target))
            .collect();
        known_types = known_types
            .into_iter()
            .map(|(name, target)| (name.to_lowercase(), target))
            .collect();
    }

    let mut namespace_constants = HashMap::new();
    if let Some(nc) = root.get("namespace_constants").and_then(|v| v.as_table()) {
        for (name, val) in nc {
            match val {
                Value::Float(f) => {
                    namespace_constants.insert(name.clone(), ConstantValue::Float(*f));
                }
                Value::Integer(i) => {
                    namespace_constants.insert(name.clone(), ConstantValue::Float(*i as f64));
                }
                Value::String(s) => {
                    if let Ok(f) = s.parse::<f64>() {
                        namespace_constants.insert(name.clone(), ConstantValue::Float(f));
                    } else {
                        namespace_constants.insert(name.clone(), ConstantValue::Str(s.clone()));
                    }
                }
                _ => {}
            }
        }
    }
    if namespaces.use_dotnet {
        for (name, value) in dotnet_namespace_constants() {
            namespace_constants
                .entry(name.clone())
                .or_insert(ConstantValue::Float(*value));
        }
    }

    // Ambient-import defaults declared via `[[esm_default]]` TOML
    // entries (Phase 4 schema). Three variants:
    //   * `kind = "named"`     — `import { name as local } from "module"`
    //   * `kind = "namespace"` — `import * as alias from "module"`
    //   * `kind = "package-root"` — Component-Model qualified-chain root
    let mut esm_defaults: Vec<EsmDefault> = Vec::new();
    if let Some(arr) = root.get("esm_default").and_then(|v| v.as_array()) {
        for entry in arr {
            let Some(tbl) = entry.as_table() else {
                continue;
            };
            let kind = tbl.get("kind").and_then(|v| v.as_str()).unwrap_or("");
            match kind {
                "named" => {
                    let Some(local) = tbl.get("local").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    let Some(module) = tbl.get("module").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    // `name` defaults to `local` when omitted.
                    let name = tbl.get("name").and_then(|v| v.as_str()).unwrap_or(local);
                    esm_defaults.push(EsmDefault::Named {
                        local: local.to_string(),
                        module: module.to_string(),
                        name: name.to_string(),
                    });
                }
                "namespace" => {
                    let Some(alias) = tbl.get("alias").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    let Some(module) = tbl.get("module").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    esm_defaults.push(EsmDefault::Namespace {
                        alias: alias.to_string(),
                        module: module.to_string(),
                    });
                }
                "package-root" | "package_root" => {
                    let Some(prefix) = tbl.get("prefix").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    let Some(module_root) = tbl.get("module_root").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    esm_defaults.push(EsmDefault::PackageRoot {
                        prefix: prefix.to_string(),
                        module_root: module_root.to_string(),
                    });
                }
                "module-export" | "module_export" => {
                    let (Some(module), Some(name), Some(target_module), Some(target_name)) = (
                        tbl.get("module").and_then(|v| v.as_str()),
                        tbl.get("name").and_then(|v| v.as_str()),
                        tbl.get("target_module").and_then(|v| v.as_str()),
                        tbl.get("target_name").and_then(|v| v.as_str()),
                    ) else {
                        continue;
                    };
                    esm_defaults.push(EsmDefault::ModuleExport {
                        module: module.to_string(),
                        name: name.to_string(),
                        target_module: target_module.to_string(),
                        target_name: target_name.to_string(),
                    });
                }
                "tree-mount" | "tree_mount" => {
                    let (Some(prefix), Some(path)) = (
                        tbl.get("prefix").and_then(|v| v.as_str()),
                        tbl.get("path").and_then(|v| v.as_str()),
                    ) else {
                        continue;
                    };
                    esm_defaults.push(EsmDefault::TreeMount {
                        prefix: prefix.to_string(),
                        path: path.to_string(),
                    });
                }
                "tree-ambient" | "tree_ambient" => {
                    let Some(path) = tbl.get("path").and_then(|v| v.as_str()) else {
                        continue;
                    };
                    esm_defaults.push(EsmDefault::TreeAmbient {
                        path: path.to_string(),
                    });
                }
                _ => {
                    eprintln!("Warning: unknown esm_default kind: {:?}", kind);
                }
            }
        }
    }

    // [bare_module_aliases] — language-specific bare→prefixed import
    // canonicalisation. Only languages that own these names route them
    // (JS routes `fs` → `node:fs`); others leave the table empty.
    let mut bare_module_aliases: HashMap<String, String> = HashMap::new();
    if let Some(tbl) = root.get("bare_module_aliases").and_then(|v| v.as_table()) {
        for (key, val) in tbl {
            if let Some(target) = val.as_str() {
                bare_module_aliases.insert(key.clone(), target.to_string());
            }
        }
    }

    Ok(LanguageProfile {
        name,
        function_return,
        result_slot_name,
        self_keyword,
        base_keyword,
        constructor_name,
        class_method_dispatch,
        dynamic_numeric_dispatch,
        enum_as_ordinals,
        case_sensitive,
        string_indexing,
        array_upper_bound_inclusive,
        negative_index_wraps,
        slice_step_zero_raises,
        tuple_literals_tagged,
        parens_for_index,
        unified_array_map,
        concat_stringifies_operands,
        entry_point,
        hoist_var,
        dynamic_add,
        function_references,
        commonjs_require,
        multi_value_tuple_returns,
        methods_bind_on_access,
        default_args_evaluated_once,
        byref_boxing,
        with_block,
        new_with_initializer,
        new_from_initializer,
        linq_queries,
        switch_fallthrough,
        throwable_is_root,
        field_hiding,
        methods_virtual_by_default,
        integer_division_on_slash,
        xor_is_logical_for_non_integers,
        for_loop_per_iteration_binding,
        bare_name_invokes_parameterless_function,
        source_function_callable_aliases,
        separate_property_method_namespace,
        reflection_type_naming,
        supports_private_fields,
        has_function_prototype_bind,
        function_invocation_members,
        has_function_constructor,
        async_wraps_body_in_try,
        ecma_error_object_shape,
        has_undefined_value,
        supports_spread_arguments,
        supports_dynamic_import,
        has_ecma_bigint,
        ecma_strict_mode,
        ecma_switch_strict_equality,
        ecma_lexical_declarations,
        has_ecma_globals,
        ecma_operator_coercion,
        abstract_equality,
        has_arguments_object,
        relaxed_call_arity,
        ecma_new_dispatch,
        ecma_array_elisions,
        has_generators,
        ecma_object_literals,
        ecma_boolean_operators,
        dynamic_member_access,
        ecma_typeof_operator,
        ecma_to_primitive,
        ecma_in_operator,
        ecma_array_method_dispatch,
        ecma_promise_methods,
        ecma_iterator_result_shape,
        member_call_on_null_error,
        string_aware_relational,
        lexical_block_scope,
        function_scoped_variables,
        unresolved_reference_throws,
        coerces_value_to_type_hint,
        ambient_this_binding,
        uses_common_resolver,
        missing_arg_is_undefined,
        materialize_bool_results,
        callable_objects,
        for_in_object_yields_keys,
        dict_literals_as_map,
        class_introspection_metadata,
        class_member_metadata,
        class_multiple_inheritance,
        supports_autoload,
        buffered_iterator_methods,
        uses_normalize_class,
        builtins,
        builtin_type_spellings: parse_builtin_types(&root),
        builtin_slots: parse_builtin_slots(&root)?,
        intrinsics,
        namespaces,
        known_types,
        value_methods,
        namespace_constants,
        array_methods,
        builtin_return_types,
        datetime_field_functions,
        esm_defaults,
        bare_module_aliases,
    })
}

#[cfg(test)]
mod builtin_slot_parse_tests {
    use super::*;
    use vybe_ast::builtin_slots::BuiltinType;
    use vybe_ast::ProtocolSlot;

    fn profile_with(section: &str) -> LanguageProfile {
        parse_profile(&format!(
            "[info]\nname = \"t\"\n\n[compiler]\n\n{section}"
        ))
        .expect("profile parses")
    }

    /// The happy path: `[builtin_slots.<type>] <slot> = "<target>"` reaches the
    /// table `Compiler::apply_builtin_slot_binding` consults.
    #[test]
    fn a_declared_override_is_parsed() {
        let p = profile_with("[builtin_slots.map]\nget_item = \"common:dart.index_get\"\n");
        assert_eq!(
            p.builtin_slots.get(BuiltinType::Map, ProtocolSlot::GetItem),
            Some("common:dart.index_get")
        );
    }

    /// A profile with no section gets an empty table, NOT a failure — that is
    /// what makes this mechanism inert for the eleven languages that declare
    /// nothing.
    #[test]
    fn a_profile_without_the_section_gets_an_empty_table() {
        let p = profile_with("");
        assert!(p.builtin_slots.is_empty());
    }

    /// A typo must be LOUD, not silently dropped.
    ///
    /// This is the discriminating test for this section. A dropped override
    /// falls through to the platform default, which is exactly the
    /// `undefined`-instead-of-`null` bug that forced `Map`/`GetItem` to be
    /// backed out on 2026-07-31 — and Dart's correctness now rests on a single
    /// entry here. Skipping-on-unknown would make a one-character typo
    /// reintroduce that bug with no signal.
    #[test]
    fn a_misspelled_type_or_slot_key_fails_the_profile() {
        for (section, bad) in [
            ("[builtin_slots.mop]\nget_item = \"common:x\"\n", "mop"),
            ("[builtin_slots.map]\nget_itm = \"common:y\"\n", "get_itm"),
        ] {
            let err = parse_profile(&format!("[info]\nname = \"t\"\n\n[compiler]\n\n{section}"))
                .expect_err("a misspelled key must fail the profile, not be skipped");
            assert!(
                err.contains(bad),
                "error must name the offending key `{bad}`, got: {err}"
            );
        }
    }

    /// Every declared override must survive parsing — the count in the TOML
    /// equals the count in the table. Asserting the count immediately after
    /// writing a profile edit is the habit that caught a `sed` matching 39
    /// entries instead of 1; this is the same check, enforced.
    #[test]
    fn every_declared_override_survives_parsing() {
        let p = profile_with(
            "[builtin_slots.map]\nget_item = \"common:a\"\nlen = \"common:b\"\n\
             [builtin_slots.string]\nget_item = \"common:c\"\n",
        );
        assert_eq!(p.builtin_slots.iter().count(), 3);
    }

    /// Slot keys are the SAME vocabulary `ProtocolSlot::from_key` round-trips,
    /// so a profile cannot name a slot the resolver has no way to look up.
    #[test]
    fn slot_keys_are_the_protocol_slot_vocabulary() {
        for slot in [ProtocolSlot::Len, ProtocolSlot::GetItem, ProtocolSlot::Eq] {
            let p = profile_with(&format!(
                "[builtin_slots.string]\n{} = \"common:z\"\n",
                slot.as_key()
            ));
            assert_eq!(
                p.builtin_slots.get(BuiltinType::String, slot),
                Some("common:z"),
                "{slot:?} did not round-trip through its key"
            );
        }
    }
}
