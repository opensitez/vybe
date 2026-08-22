# Directives — declaring language policy in the AST

**The rule: everything is expressed by the AST.**

A profile property is not a lighter-weight alternative to an AST node. It is a
worse one, always, and the number of things that read it is irrelevant. The same
goes for a gate, a language-name check, a per-language table in a walker, a
thread-local side map, and an inference pass. All six are the same mistake with
different spelling: **a fact about the program stored somewhere the program
cannot see.**

This document exists so that argument does not have to be made from scratch
again.

---

## 1. Why not a profile flag

### It hides bugs

A property must be consulted at every site that needs it, and sites forget.

- **`fold_case`** used to be a `case_sensitive` flag on the compiler, so all 33
  call sites had to remember to write `!self.case_sensitive &&` themselves.
  **23 of them did not.** Two silently broke Go: a local `ab` matched the class
  `AB`, so `ab.Get()` stopped resolving as a class reference and compiled to a
  call that passed no receiver. It is now `Scope.fold_case`, a property of the
  scope that does the resolving — and `Scope::new` *requires* it rather than
  defaulting it, precisely so a missed construction site is a build error
  instead of a silent mis-resolution.

- **`use_dotnet`** is the live example of a gate. It made **PHP quirks reachable
  only when a .NET flag was set** — behaviour that existed, was correct, and did
  nothing, because an unrelated switch was off. That is the signature failure
  mode of a gate: not a wrong answer, a missing one, with no error anywhere.

### It is not readable for reflection

The AST is available to the running program. Profile data never is.

A fact placed in a profile row is invisible to reflection **permanently**. PHP's
`ReflectionParameter::isPassedByReference()` is a real API, and `getParameters()`
already has tests in this tree. If "this parameter is by reference" lives in a
profile row, reflection cannot report it — not because nobody wired it up, but
because the information is not in the tree the program can read.

Ask of any fact: *if a program asked for this through reflection, could it get an
answer?* If the answer is no, the carrier is wrong.

### A flag with one language IS a language check

`args_pass_by_reference` is read only by Fortran. `promote_addr_taken_at_entry`
is read only by C. Nobody wrote `profile.name == "fortran"`, but the flag has one
consumer and one language, so it is that check with a nicer name. There used to
be **hundreds** of language checks and they hid obscure bugs; the count going
down is the point. Arguing that any individual one is harmless because few things
read it is how the count stops going down.

---

## 2. The carriers that exist

All of them are AST. The choice is never "AST or profile" — it is *which node*.

| carrier | granularity | for |
|---|---|---|
| `Module.directives: Directives` | whole file | the walker states the language's default |
| `StmtKind::Directive { set, scope }` | `DirectiveScope::Block` or `Module` | an in-source change from that point |
| `StmtKind::ScopeDecl { kind, names }` | positional, one scope | `global $x`, `nonlocal x` |
| `Modifiers` on a declaration | one declared thing | `is_static`, `is_readonly`, `is_extension`, `visibility` |
| a field on the node itself | one node | `Param.pass_by`, `Local.holds_reference` |

### `Directives`

```rust
pub struct Directives {
    pub array_storage: Option<ValueStorage>,
    pub reference_binding: Option<PassBy>,
    pub set_semantics: Option<SetSemantics>,
    pub receiver_binding: Option<ReceiverBinding>,
    pub app_shell: Option<AppShell>,
    pub shift_overflow: Option<ShiftOverflow>,
}
```

**Every field is an `Option`, and that is load-bearing.** One type serves both
carriers: a module's declared defaults, and an in-source delta that changes
exactly one thing. They combine only through `overlay` — a `Some` wins, a `None`
inherits. There is no other way to merge them, so "what is in force here" always
has one answer.

This is not the finished set. It is what has been moved so far.

`shift_overflow` is the worked example of §3 question 1 done right, and of the
question next to it answered differently. wasm MASKS a shift count, so `1 << 32`
is `1`; Fortran's `ISHFT` yields ZERO once the count reaches `BIT_SIZE`, and
`gfortran` proves it. Nothing in either operand's type distinguishes those — the
language does, so it is lexical, so it is a directive.

The *width* of a bit operation is the opposite call. It looks like the same kind
of fact, and it is not: `integer(kind=8)`, `uint32` vs `uint64` and
`Integer.bitCount` vs `Long.bitCount` are all DECLARED types, which is question
2/3. It lives on the node as `BitLane`, the way wasm spells `i32.popcnt` and
`i64.popcnt` as different instructions. Putting it in a directive would have
given one fact two homes — the failure §3's worked disambiguation exists to
prevent.

### `DirectiveScope` — why two

Real languages have both, so the AST declares which one it means:

- **`Block`** restores at the end of the enclosing block. JS `"use strict"` at a
  function head; a pragma scoped to a nested block.
- **`Module`** survives every block end. Pascal `{$R+}` switched on halfway down
  a procedure stays on for the rest of the unit; C `#pragma`; PHP `ini_set`.

The implementation is the whole trick, and it is four lines
(`primitives/statements.rs::apply_directive`):

```rust
DirectiveScope::Block  => if let Some(top) = self.directives.last_mut() { top.overlay(set) },
DirectiveScope::Module => for frame in &mut self.directives { frame.overlay(set) },
```

`Block` overlays the innermost frame, so exiting the block pops it away.
`Module` **writes through every frame**, so popping cannot lose it. That is the
entire difference between "restores on exit" and "outlives the block."

### Cost

`Compiler.directives: Vec<Directives>`, innermost last, never empty. Frame 0 is
installed from `module.directives`. A block pushes a frame **only if its body
actually contains a `Directive` statement** (`stmts_have_directive`, a shallow
scan — a nested block runs its own). So ordinary code pays one linear scan per
block and nothing else. Module top level needs no frame: frame 0 *is* the
enclosing scope.

Wired at the same places as the existing `in_strict` save/restore, which is the
precedent to follow if a fourth is needed: `StmtKind::Block`, the try body, and
`compile_function_decl`.

---

## 3. Which carrier — the test

Three questions, in order. The first `yes` is the answer.

1. **Does it govern a region of CODE?** → a **directive**.
   It is lexical. It applies to whatever is written inside it, and moving code
   changes what it obeys. `Option Explicit`, `declare(strict_types=1)`,
   `{$R+}`, `"use strict"`, `ini_set`, "arrays copy on assignment in this file",
   "references bind rather than store here".

2. **Does it travel with a VALUE?** → a **stamp or slot** on the instance.
   It moves with the data across function and even language boundaries. A Pascal
   record assigned inside PHP copies because its stamp came with it.

3. **Does it describe a DECLARED thing?** → a field or `Modifiers` **on that
   declaration**.
   How a callee takes its arguments; whether a member is static, readonly, an
   extension. It is true wherever the thing is used, so it belongs to the thing.

**A worked disambiguation.** "PHP's `bindParam` takes argument 1 by reference" is
question 3, not question 1: it is true regardless of which file calls it or what
`ini_set` ran. It travels with the callee. Meanwhile PHP's historical
`allow_call_time_pass_reference` — may the *caller* write `&` at the call site? —
is question 1, because it governs the code doing the calling.

Same subject area, different questions, different carriers. Getting this wrong is
how one fact ends up with two homes.

---

## 4. The worked example: `ScopeDecl`

This is the conversion to copy, because it is done and it is in the tree.

**Before.** PHP function bodies do not see module globals without `global $x`.
Every other language chains outward. That single difference was spread across
**five `profile.name == "php"` checks** and a `php_function_globals` field on the
shared compiler.

**After.** Two pieces, both AST:

```rust
// vybe_ast — the language DECLARES the policy
StmtKind::ScopeDecl { kind: ScopeDeclKind, names: Vec<String> }
enum ScopeDeclKind { Closed, Global }

// the shared machinery gains a real PROPERTY
enum ScopeResolution { Chain, Closed }
struct Scope { resolution: ScopeResolution, open_names: HashSet<String>, … }
```

`scope.rs` states the principle in its own doc comment:

> *"This is a property of the SCOPE, not of a language."*

`Closed` carries `names` for the exceptions — PHP's superglobals are visible
everywhere without being imported, and **that list is PHP's to supply**, not
shared code's to know. Any language wanting closed function scopes gets it by
declaring it. No shared code learned the word "php".

---

## 5. Anti-patterns, with live examples

Each of these is a fact stored where the program cannot read it.

| shape | example in this tree |
|---|---|
| **profile behaviour flag** | `args_pass_by_reference` (Fortran-only), `promote_addr_taken_at_entry` (C-only) |
| **gate** | `use_dotnet` — 24 refs, 14 in shared crates |
| **language-name check** | 13 left in `vybe_compiler` + `vybe_runtime` |
| **name table in a walker** | a hardcoded list of method names the walker treats specially |
| **thread-local side map** | Kotlin's `CLASS_PROPERTIES: name → readonly`, carrying whether a primary-constructor `val` param is read-only |
| **reusing a field as a marker** | C# writes `PassBy::Const` on a parameter to mean "extension receiver" — the comment says *"Reuse `Const` as an internal marker"* |
| **inference** | Fortran marks a parameter `Alias` only if it can *see* the body mutate it |

The last two deserve emphasis.

**Reusing a field as a marker** makes one field answer several unrelated
questions. `PassBy::Const` currently means four things at once: COBOL
`BY CONTENT` (copy in, never out), Fortran `intent(in)` (read-only), Pascal
`const`/`constref` (read-only, and `constref` is explicitly *by-ref* — proving
mechanism and read-only are independent), and C#'s extension receiver (not a
passing mode at all). Any change to the enum silently breaks whichever meaning
you were not thinking about.

**Inference is the deepest one**, because it looks like cleverness. Deriving a
declared property by scanning code answers "what does this program appear to do"
when the question was "what did the author declare". It is wrong whenever the
scan's assumptions do not hold, and it fails silently, and the fix is always to
declare the thing.

### The tell

**Several languages each inventing a different channel for the same fact means
the AST is missing a field.** Five languages need declaration facts on a
parameter — C# `this` and `in`, Kotlin `val`/`var`, Pascal `const`, Fortran
`intent(in)` — and produced five different mechanisms, because `Param` offers
none. That is not five local quirks. It is one missing `Param.modifiers`.

---

## 6. Before adding anything: SEARCH

Most "missing" carriers already exist. Adding a second one is how the tree got
into this state.

- `Modifiers.is_extension` **already exists** and is used by Kotlin, VB (which
  both sets it from `<Extension>` and reads it) and Pascal. C# ignored it and
  reused `PassBy::Const`. That fix needs **no shared change at all**.
- `Modifiers.is_readonly` **already exists**, which is very likely where Pascal
  `const` and Fortran `intent(in)` belong.

So: find how the other languages already express it. If nine of them do it
through a mechanism and one does not, the odd one out is the bug — not evidence
that shared code needs a new column.

---

## 7. Adding a directive

1. **Search first** (§6). Confirm no carrier exists.
2. **Answer §3's three questions.** If it is not lexical, it is not a directive —
   put it on the declaration or the value instead.
3. **Add an `Option<T>` field to `Directives`.** Never a bare `bool`: `None` must
   mean "not stated" and remain distinguishable from "stated false", or `overlay`
   cannot inherit correctly.
4. **Extend `overlay`** — it is hand-written on purpose, so each field's
   combination rule is a decision someone made rather than a derive.
5. **The walker sets the language default** on `Module.directives`, and maps the
   language's own in-source syntax to `StmtKind::Directive`, choosing `Block` or
   `Module` scope to match what that syntax really does.
6. **Shared code reads the directive.** Never a language name, never a profile
   flag.
7. **Delete the old channel in the same change.** A directive added beside a
   surviving flag is worse than either alone — now the fact has two homes and
   they will drift.

### Landing it without breaking the build

Adding a required field to a shared struct with exhaustive literals breaks every
one of them at once. A required field on `CtorSpec` did exactly that on
2026-08-07 and broke the build for the user *and* another agent working the same
tree. `Param` today has **251 exhaustive literals** across `languages/*`,
`crates/vybe_compiler` and `platforms/*`, and no `Default` derive.

Sequence it so no step breaks anything:

1. `#[derive(Default)]` on the struct — pure addition.
2. Migrate literals to `..Default::default()`, **per file** — still non-breaking,
   and a collision with another agent stays recoverable.
3. Only then add the field.

And prefer one field that absorbs the whole category over several ad-hoc bools.
`Param.modifiers: Modifiers` costs the same 251 sites as two bools would, and
pays for `readonly`, `extension`, and primary-constructor visibility (Kotlin
`val`, C# 12) at once instead of returning for each.

---

## 8. Checklist

Before writing a special case, in order:

- [ ] Have I **searched** for an existing carrier? (`Modifiers`, `Directives`,
      an existing node field, a slot)
- [ ] Do **other languages** already express this? How? Am I the odd one out?
- [ ] Which of §3's **three questions** is this — code region, value, or
      declaration?
- [ ] Could a program read this fact through **reflection** once I am done?
- [ ] Am I adding a **second channel** for something that already has one?
- [ ] Am I **inferring** something the author could have declared?
- [ ] Does my change let me **delete** a flag, a gate, or a language check?

If the last box stays unticked, the change is probably adding to the problem.

---

## 9. Unification is not about reuse — it is what makes languages interoperate

The weakest argument for unifying a mechanism is "less duplicated code". That is
a side effect. There are two real reasons, and both are worth more than the code
saved.

### 9.1 It is the only thing that makes cross-language programs work

A PHP class inheriting a Python class is not a trick. It is what falls out of
both languages landing on **one** class model — and it stops working the moment
either language keeps its own private lowering.

The concrete machinery: 14 languages ship a `languages/*/src/protocol.rs` that
maps their own spelling onto a shared canonical member.

| language | spelling | canonical |
|---|---|---|
| PHP | `__toString` | `tostring` / `ToString` |
| Python | `__str__` | `tostring` / `ToString` |
| PHP | `__clone` | `clone` / `Clone` |
| Python | `__copy__` | `clone` / `Clone` |
| PHP | `__get` / `__set` | `getattr` / `setattr` |

Because the *spelling* is per-language and the *slot* is shared, a Python class
that defines `__str__` is printable from PHP, and PHP cloning a Python object is
just the slot. Nobody wrote a php↔python bridge. There is nothing to bridge:
both walkers normalized to the same member, and inheritance is ordinary member
lookup over `__types`.

**The counter-example proves it.** Classes cross language boundaries in this tree.
Enums did not — because enum lowering stayed per-language while classes went
through the shared model. Same program, same runtime, and the difference in
whether interop works is precisely the difference in whether the concept was
unified. A per-language mechanism is a wall between languages, whether or not
anyone intended one.

So: **a fact expressed once in the AST is automatically available to all 17
languages. A fact expressed in a walker, a profile row, or a language-name branch
is available to exactly one, forever.** That is the whole argument for the AST,
restated as a capability rather than a style preference.

### 9.2 Centralizing collapses whole classes of obscure bugs

When N spellings exist for one concept, they **drift**, and the drift is silent
because each spelling looks locally correct. Every example below is a real bug
from this tree, and every one is the same bug:

| duplicated thing | how it failed |
|---|---|
| `case_sensitive` at 33 call sites | 23 omitted it; Go's `ab` matched class `AB` and `ab.Get()` compiled with no receiver |
| two spellings of "address of a name" | local-then-global resolution had drifted between them; unifying to one resolver fixed both |
| Go's private copy of the place converter | deleted in favour of the shared one; slice unchanged at 99/12 — the duplicate was pure risk |
| six reference spellings | none could express "alias" vs "copy-in/copy-out", so the distinction was made by inference instead |

And centralizing pays forward, not just backward. Two from one session:

- Fixing where a bound parameter is read fixed **mysqli for free**, because PDO
  and mysqli already shared one `__bound_params` key. One fix, two adapters,
  because there was one place to fix.
- Making PHP's binder emit the *same* `&$n` a programmer could write meant the
  module-wide address-taken scan and pointer-cell promotion picked it up with
  **no new mechanism at all** — the existing machinery already handled that node.

That is the compounding return: each thing unified makes the *next* fix land in
one place instead of N, and makes it reach every language at once.

### 9.3 The test

> If I fix this bug, does the fix reach every language that has this concept —
> or only the one I am looking at?

If only one, you are not fixing a bug; you are adding the N+1th spelling. And
ask the interop question explicitly, because it is the one that gets forgotten:

> If a class from another language were passed in here, would this still work?

`profile.name == "php"` always answers no.

## 10. Where each layer's answer lives — and why "JS-shaped" is not "JS-limited"

The AST is JS-shaped and the host is ECMA. That is a starting point, not a
ceiling, and misreading it is the source of two opposite mistakes: *"ECMA has no
such method, so my language needs its own path"*, and *"the AST has no such node,
so this belongs in my walker"*. Both are wrong, and both produce the N+1th
spelling.

### 10.1 `primitives/*` — the mechanics ECMA does not provide

`crates/vybe_compiler/src/primitives/` is **90 modules**, and the size is the
point: it is everything real languages need that ECMA does not hand you.

```
references pointers addressable_storage heap memory      records tuples sets dict
channels threading async_ops generators                  enums enum_lowering generics
sprintf packing codepoints string_encoding csv xml       overloads dispatch delegates
case_insensitive_collections sorted_collection           namespaces reflection metadata
```

ECMA has no COBOL `PIC` packing, no Go channel rendezvous, no Fortran
multi-dimensional slice, no Pascal set type, no case-insensitive collection, no
`sprintf`, no by-reference cell. `primitives/*` owns them **once**, in the shared
compiler, for every language that asks.

The rule that follows:

> A method your language needs and ECMA lacks is a **`primitives/*` gap**, not a
> licence to write a private path in your walker or an adapter.

This is why "no stdlib gap" is a rule rather than a slogan. There is no category
of "things ECMA doesn't do, so we each do our own" — that category is
`primitives/*`, and the second one language writes its own, the concept has two
implementations that will drift (§9.2). Adapters are for a language's *surface*
mapping onto a primitive; they are not a place to reimplement one.

And it composes: `references.rs` gave PHP's `bindParam` its reference for free,
because C, Go, Pascal and C# had already paid for the cell machinery. The tenth
language to need a mechanism should write nearly nothing.

### 10.2 `vybe_ast` — the semantics neither JS nor ECMA has

The AST is JS-*shaped* because a C-family expression/statement tree is a good
substrate. It is not JS-*limited*. It already carries concepts JS has never had:

| node | concept | from |
|---|---|---|
| `PassBy` | `Value` / `Ref` / `Alias` / `Out` / `Const` | Pascal `var`, C# `ref`/`out`, PHP `&`, Fortran `intent`, COBOL `BY REFERENCE` |
| `RefOf(PlaceExpr)` | a reference to a STORAGE LOCATION | C, Go, Pascal, PHP |
| `Chan(ChanOp)` | channel send/receive/select | Go, Rust, Kotlin |
| `ScopeDecl` | scope resolution policy | PHP `global`, Python `nonlocal` |
| `Directive` / `Directives` | lexical language policy | `Option Explicit`, `{$R+}`, `strict_types` |
| `ProtocolSlot` | a language-neutral member identity | 14 `protocol.rs` files |
| `ValueSemantics` | value vs reference storage for a declared type | Pascal records, C# structs |
| `MemoryDecl` | linear memory and data segments | WAT |
| `MatchStatement` | full pattern matching | Python, Rust-style |

None of these is "a JS feature bent into shape". Each is a **semantic fact some
real language declares**, given one representation so that every language
declaring the same fact says it the same way.

So the second rule:

> The AST does not stop at what JS can express. It stops at what no language
> means. If your language declares something the AST cannot say, the answer is a
> node — not a walker workaround.

### 10.3 Normalization: syntax down, semantics up

This is the division that makes the whole thing hold:

| layer | owns | must never |
|---|---|---|
| **walker** | SYNTAX. Spelling, sugar, argument order, statement forms | encode semantics in a private table, side map, or inference |
| **`vybe_ast`** | SEMANTICS. What the author declared, in one vocabulary | carry a language name, or a field only one language reads |
| **`primitives/*`** | MECHANICS. How a declared fact is emitted | branch on which language it is compiling |
| **`platforms/ecma`** | ECMA-262 conformance | acquire a mode for one language's convenience |

**Normalization is a lossless rewrite of spelling into vocabulary.** Kotlin's
`data class`, Python's `@dataclass` and Pascal's `record` are three syntaxes for
one declared fact, so all three walkers emit the same declaration and one shared
path derives the members. Dart's `length`, Python's `__len__` and PHP's `count`
are three spellings of one slot.

Two failure modes, both common:

- **Under-normalizing** — the walker passes its own syntax through and shared
  code grows a branch to cope. That branch is a language check.
- **Over-normalizing** — the walker *decides* something instead of declaring it,
  and the decision is invisible downstream. Fortran marking a parameter as
  aliased only when it can *see* the body mutate it is this: the language rule
  ("dummy arguments are by reference") became a guess about one function body.

The test for a walker change:

> Am I rewriting how this is *spelled*, or deciding what it *means*?

Spelling is yours. Meaning belongs in the AST, where every language and the
program's own reflection can read it.

## 11. Summary

- **Everything is expressed by the AST.** Profile properties and gates hide bugs
  and are invisible to reflection.
- **"Few things read it" is not a defence.** A flag with one language is a
  language check.
- **Both granularities already exist** — `Module.directives` for a file,
  `StmtKind::Directive { scope: Block | Module }` for a region, and `Module`
  scope writes through every frame so it can outlive its block.
- **Pick the carrier by what the fact belongs to:** code region → directive,
  value → stamp, declared thing → declaration.
- **Search before adding.** `Modifiers.is_extension` and `Modifiers.is_readonly`
  are already there.
- **Delete the old channel in the same change**, or the fact now has two homes.
- **Unification is not about reuse.** A fact expressed once in the AST is
  available to all 17 languages and lets them interoperate — a PHP class can
  inherit a Python one because both normalize to the same class model and the
  same protocol slots. A fact expressed in a walker, a profile row or a
  language-name branch is available to exactly one language, forever, and is a
  wall between languages whether or not anyone meant it to be.
- **N spellings drift silently.** Every duplicated mechanism in this tree
  eventually diverged, and each copy looked locally correct while it did.
- **JS-shaped is not JS-limited.** ECMA lacking a method means a `primitives/*`
  gap (90 modules already: channels, records, packing, references, sets…), not a
  private path in your walker. The AST lacking a node means the AST needs the
  node — it already carries `PassBy`, `RefOf`, `Chan`, `ScopeDecl`,
  `ProtocolSlot`, `MemoryDecl`, none of which JS has.
- **Walkers own SPELLING; the AST owns MEANING.** If your walker is *deciding*
  rather than *rewriting*, the decision belongs in the tree.

Related: `referenceplan.md` (§10g–§10i for the parameter-modifier case),
`proxyplan.md`, `flexclassplan.md`.
