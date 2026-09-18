Here is a comprehensive breakdown of **150 brand-new, semantically distinct PowerShell core language categories** designed to maximize language coverage. 

Each category targets core language and engine features that are **not covered by the 388 existing test directories** (0% overlap, no external OS binaries, pure PowerShell semantics):

---

### 1. Advanced Cmdlet & Function Lifecycle (12 Categories)
1. `process_record_streaming` — Streaming execution semantics through the `Process` block vs batched pipeline collections.
2. `clean_block_semantics` — PowerShell 7.3+ `clean { }` block execution guarantees during pipeline interruptions, errors, and early termination.
3. `dynamic_parameter_runtime_generation` — Runtime construction of parameters using `RuntimeDefinedParameterDictionary` and `RuntimeDefinedParameter`.
4. `pscmdlet_write_debug_record` — Emitting typed debug records with invocation context and preference handling via `$PSCmdlet.WriteDebug()`.
5. `pscmdlet_write_information_record` — Creating and writing tagged `InformationRecord` objects via `$PSCmdlet.WriteInformation()`.
6. `pscmdlet_write_warning_record` — Writing warning records with custom attributes and preference actions.
7. `pscmdlet_write_error_record` — Constructing structured non-terminating errors with `ErrorCategory`, `TargetObject`, and `RecommendedAction`.
8. `pscmdlet_transaction_binding` — `[CmdletBinding(UseTransaction=$true)]` and `$PSCmdlet.CurrentPSTransaction` lifecycle.
9. `supports_paging_parameters` — `[CmdletBinding(SupportsPaging=$true)]` and paging engine hooks (`$PSCmdlet.PagingParameters.First/Skip`).
10. `supports_wildcards_attribute` — `[SupportsWildcards()]` parameter attribute resolution on string patterns.
11. `credential_attribute` — `[Credential()]` attribute type-coercion on `PSCredential` parameters.
12. `alias_attribute_multiple` — Multiple `[Alias()]` attributes on single parameters with priority resolution.

---

### 2. AST, Parsing, & Compiler Architecture (13 Categories)
13. `ast_safe
<truncated 14577 bytes>
nrelated_types` — Type-specific catch blocks: `catch [System.IO.IOException], [System.TimeoutException]`.
136. `finally_execution_on_pipeline_break` — Guaranteeing `finally` execution when a downstream pipeline halts with `break`.
137. `finally_execution_on_return` — `finally` execution order when a `return` statement is encountered in `try`.
138. `error_variable_capture_detailed` — Appending error records to custom variables using `-ErrorVariable +myErr`.
139. `error_invocation_info_script_bounds` — Accurate script boundary reporting in `$_.InvocationInfo`.
140. `error_record_category_activity` — Custom `ErrorCategoryInfo.Activity` metadata validation on errors.

---

### 12. Language Operators & Built-in Semantics (10 Categories)
141. `operator_regex_options_inline` — Inline regex flags (`(?i)`, `(?m)`, `(?s)`) with `-match` and `-replace`.
142. `operator_join_single_operand` — Prefix array join semantics: `-join $collection`.
143. `operator_split_max_substrings` — Splitting strings with bounded chunk counts (`$str -split ',', 3`).
144. `operator_is_isnot_custom_types` — Type verification using `-is` and `-isnot` against generic and interface types.
145. `operator_as_safe_coercion` — Safe type casting with `-as` returning `$null` on conversion failures without throwing.
146. `operator_shl_shr_boundary_overflow` — 32-bit and 64-bit boundary overflow handling with `-shl` and `-shr`.
147. `operator_xor_bitwise_and_boolean` — Boolean `-xor` logical evaluation vs integral `-bxor`.
148. `psdefaultparameter_values_patterns` — Target cmdlet wildcard matching in `$PSDefaultParameterValues` (`'*-Object:Property'`).
149. `psdefaultparameter_values_scriptblock` — Dynamic parameter default values computed via ScriptBlocks.
150. `automatic_variable_matches_scope` — `$Matches` automatic variable scoping, capture groups, and lifecycle.

---

Every category in this list is guaranteed unique against the repository's 388 existing categories and focuses strictly on deep PowerShell language features.
