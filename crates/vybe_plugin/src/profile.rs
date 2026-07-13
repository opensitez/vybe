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
    /// via `vybe_emitter::tuples::TUPLE_TAG`) rather than a plain list, so
    /// `repr`/`type()`/slicing tell a tuple from a list. Python opts in;
    /// languages whose `(a, b)` is just grouping/an array keep this `false`.
    pub tuple_literals_tagged: bool,

    /// Whether parens are used for both calls and indexing (VB: arr(i)).
    pub parens_for_index: bool,

    /// Entry point function name to auto-call if defined (e.g. "main").
    pub entry_point: Option<String>,

    /// JS: `var` declarations are hoisted to function scope.
    pub hoist_var: bool,

    /// JS: `+` operator uses dynamic add (string concat if either operand is string).
    pub dynamic_add: bool,

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
    /// (`compiler::resolver`) instead of the legacy hardcoded arms.
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
    /// Use .NET BCL resolution from crate::platforms::dotnet::emitter.
    /// When true, the compiler uses dotnet::namespace_roots(), dotnet::default_interface_imports(),
    /// dotnet::resolve_dotted_name(), etc. — the full .NET resolution pipeline.
    pub use_dotnet: bool,
    /// When true, uses the shared .NET dotted-name resolver
    /// (`resolve_dotted_name`) for `Foo.Bar(...)` call sites. VB
    /// needs this because `Thread.Sleep` / `String.Format` etc. must
    /// route to host imports. C# generally doesn't — the default
    /// member-call path handles static dispatch on user classes and
    /// the dotnet-class ctor path handles `new Form()`-style uses.
    /// Splitting this flag from `use_dotnet` lets C# install Form /
    /// Button / Point / Size as callable globals **without** pulling
    /// the eager import-prefix fallback that mis-routes user static
    /// calls (`MathUtils.Fact(5)` → `system.mathutils.fact`).
    pub use_dotnet_resolver: bool,
    /// Additional imports beyond the defaults (e.g. "microsoft.visualbasic" for VB).
    pub extra_imports: Vec<String>,
    /// Known namespace roots (used when use_dotnet is false).
    pub roots: Vec<String>,
    /// Default imports always available (used when use_dotnet is false).
    pub default_imports: Vec<String>,
    /// Known constants (property access, not function call).
    pub constants: Vec<String>,
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
    let commonjs_require = compiler
        .get("commonjs_require")
        .and_then(|v| v.as_bool())
        .unwrap_or(false);
    let multi_value_tuple_returns = compiler
        .get("multi_value_tuple_returns")
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
    let string_aware_relational = compiler
        .get("string_aware_relational")
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
                    if let Some(emit) = parse_emit(emit_str) {
                        map.insert(
                            name.clone(),
                            BuiltinDef {
                                emit,
                                min_args,
                                max_args,
                            },
                        );
                    }
                }
            }
        }
        map
    }

    fn parse_emit(s: &str) -> Option<BuiltinEmit> {
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
                            if let Some(emit) = parse_emit(emit_str) {
                                map.entry(name.clone()).or_default().push(BuiltinDef {
                                    emit,
                                    min_args,
                                    max_args,
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
                    if let Some(emit) = parse_emit(emit_str) {
                        map.entry(name.clone()).or_default().push(BuiltinDef {
                            emit,
                            min_args,
                            max_args,
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
    let mut array_methods = parse_string_table(&root, "array_methods");

    let namespaces = if let Some(ns) = root.get("namespaces") {
        NamespaceConfig {
            use_dotnet: ns
                .get("use_dotnet")
                .and_then(|v| v.as_bool())
                .unwrap_or(false),
            use_dotnet_resolver: ns
                .get("use_dotnet_resolver")
                .and_then(|v| v.as_bool())
                // Default: the resolver is on whenever `use_dotnet` is —
                // that's how VB works today. Languages that want the
                // class registration but NOT the eager dotted-name
                // rewrite (C#) can override to `false`.
                .unwrap_or_else(|| {
                    ns.get("use_dotnet")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false)
                }),
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
        entry_point,
        hoist_var,
        dynamic_add,
        commonjs_require,
        multi_value_tuple_returns,
        byref_boxing,
        with_block,
        new_with_initializer,
        new_from_initializer,
        linq_queries,
        switch_fallthrough,
        throwable_is_root,
        field_hiding,
        separate_property_method_namespace,
        reflection_type_naming,
        supports_private_fields,
        has_function_prototype_bind,
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
        unresolved_reference_throws,
        coerces_value_to_type_hint,
        ambient_this_binding,
        uses_common_resolver,
        missing_arg_is_undefined,
        materialize_bool_results,
        supports_autoload,
        buffered_iterator_methods,
        uses_normalize_class,
        builtins,
        intrinsics,
        namespaces,
        known_types,
        value_methods,
        namespace_constants,
        array_methods,
        esm_defaults,
        bare_module_aliases,
    })
}
