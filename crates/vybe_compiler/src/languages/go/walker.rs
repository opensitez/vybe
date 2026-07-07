//! Go walker — pest `Pair<Rule>` → `vybe_compiler::ast::Module`.
//!
//! Walks the parse tree produced by `grammar.pest` into the common AST.
//!
//! ## Go-specific normalisations
//!
//! - **Multiple return values**: Go functions can return multiple values.
//!   For simplicity we compile to returning a single array/tuple.
//! - **Short variable declaration** (`:=`): Maps to `VarDecl` with `Let`.
//! - **Methods**: Go methods on structs are compiled into `StructDecl`
//!   fragments with the receiver kept as the first explicit parameter.
//! - **Structs**: Mapped to `StructDecl` with fields.
//! - **Interfaces**: Mapped to `InterfaceDecl`.
//! - **`range`**: Mapped to `ForIn` with `of: true`.
//! - **`defer`**: Lowered to a per-function stack of zero-arg closures that
//!   drain from a synthesized `finally` block in LIFO order.
//! - **`go`**: Lowered to the shared thread/task emitter surface.
//! - **`fallthrough`**: Not yet supported in switch.
//! - **`select`**: Lowered as a compile-safe block for the dummy concurrency tests.
//! - **`chan` / `<-`**: Lowered into compile-safe object/array operations.
//! - **`nil`**: Mapped to `ExprKind::Lit(Literal::Null)`.
//! - **`make` / `new`**: `make` for slices/maps is rewritten to array/dict
//!   creation. `new(T)` becomes `&T{}` (pointer to zero value).
//! - **`append`**: Rewritten to slice concat so the updated slice value is preserved.
//! - **`len` / `cap`**: Builtin functions mapped to host calls.
//! - **`panic` / `recover`**: Mapped to throw/try-catch.
//! - **`_` blank identifier**: Ignored in assignments.

use super::{GoParser, Rule};
use crate::ast::*;
use crate::common::channels;
use pest::Parser;
use pest::iterators::Pair;
use std::collections::{HashMap, HashSet};

// ══════════════════════════════════════════════════════════════════════════════════════════
// Entry point
// ══════════════════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let (package_name, mut body, imports) = walk_go_source(source)?;

    // Inject Go-source runtime preludes (small plain-Go helper libraries that
    // compile through the same pipeline — no adapter bytecode, no host fns)
    // when the program uses them.
    let mut prelude: Vec<Statement> = Vec::new();
    if go_uses_errors_runtime(source) {
        prelude.extend(go_prelude_body(GO_ERRORS_PRELUDE)?);
    }
    if source.contains("sort.") {
        prelude.extend(go_prelude_body(GO_SORT_PRELUDE)?);
    }
    if source.contains("strings.") {
        prelude.extend(go_prelude_body(GO_STRINGS_PRELUDE)?);
    }
    if source.contains("strconv.") {
        prelude.extend(go_prelude_body(GO_STRCONV_PRELUDE)?);
    }
    if source.contains("time.") {
        prelude.extend(go_prelude_body(GO_TIME_PRELUDE)?);
    }
    if source.contains("net/url") {
        prelude.extend(go_prelude_body(GO_NETURL_PRELUDE)?);
    }
    if source.contains("atomic.") {
        prelude.extend(go_prelude_body(GO_ATOMIC_PRELUDE)?);
    }
    if source.contains("slices.") || source.contains("maps.") {
        prelude.extend(go_prelude_body(GO_SLICES_MAPS_PRELUDE)?);
    }
    // slog handlers write to a bytes.Buffer, so slog pulls in the bytes prelude.
    if source.contains("bytes.") || source.contains("slog.") {
        prelude.extend(go_prelude_body(GO_BYTES_PRELUDE)?);
    }
    if source.contains("slog.") {
        prelude.extend(go_prelude_body(GO_SLOG_PRELUDE)?);
    }
    if !prelude.is_empty() {
        prelude.append(&mut body);
        body = prelude;
    }

    Ok(normalize_go_module(Module {
        name: package_name,
        language: Lang::Go,
        body,
        imports,
    }))
}

/// Whether the source references the errors/Errorf runtime surface handled by
/// the injected prelude. Cheap textual gate so ordinary programs don't pay for
/// the helper functions.
fn go_uses_errors_runtime(source: &str) -> bool {
    source.contains("errors.") || source.contains("Errorf")
}

/// Walk a prelude source and return its top-level statements, dropping the
/// placeholder `main` used to keep the snippet a complete program.
fn go_prelude_body(source: &str) -> Result<Vec<Statement>, String> {
    use std::collections::HashMap;
    use std::sync::{Mutex, OnceLock};
    // Each prelude constant is fixed source, re-walked on every compile; cache
    // the parsed statements by content (same pattern as the profile cache) and
    // hand out clones so each caller gets its own mutable copy.
    static CACHE: OnceLock<Mutex<HashMap<String, Vec<Statement>>>> = OnceLock::new();
    let cache = CACHE.get_or_init(|| Mutex::new(HashMap::new()));
    if let Some(hit) = cache.lock().unwrap().get(source) {
        return Ok(hit.clone());
    }
    let (_, body, _) = walk_go_source(source)?;
    let stmts: Vec<Statement> = body
        .into_iter()
        .filter(|stmt| !matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "main"))
        .collect();
    cache.lock().unwrap().insert(source.to_string(), stmts.clone());
    Ok(stmts)
}

/// Go-source runtime prelude for the closure-based `sort` package helpers.
/// `sort.Search`, `sort.Slice`/`SliceStable`, `sort.SliceIsSorted`, and the
/// `*AreSorted` helpers are rewritten in the walker to call these (they need a
/// closure over the target slice, so the type-specific part is synthesized at
/// the call site). Insertion sort keeps `SliceStable` stable.
const GO_SORT_PRELUDE: &str = r#"package main

func __go_sort_search(n int, f func(int) bool) int {
	lo, hi := 0, n
	for lo < hi {
		mid := (lo + hi) >> 1
		if f(mid) {
			hi = mid
		} else {
			lo = mid + 1
		}
	}
	return lo
}

func __go_sort_slice(n int, less func(int, int) bool, swap func(int, int)) {
	for i := 1; i < n; i++ {
		for j := i; j > 0 && less(j, j-1); j-- {
			swap(j, j-1)
		}
	}
}

func __go_sort_is_sorted(n int, less func(int, int) bool) bool {
	for i := 1; i < n; i++ {
		if less(i, i-1) {
			return false
		}
	}
	return true
}

func main() {}
"#;

/// Go-source runtime prelude for the `time` package. `time.Time` is modeled as
/// `{sec, nsec, loc}` — the wasi wall-clock datetime shape — for nanosecond
/// precision. `Now()` reads `wasi:clocks/wall-clock`; calendar breakdown uses
/// `ecma:date` (both reached via the `__go_date_*` / `__go_wall_now` builtins).
const GO_TIME_PRELUDE: &str = r#"package main

type __goLoc struct {
	name   string
	offset int
}

type __goTime struct {
	sec  int
	nsec int
	loc  __goLoc
}

func __go_time_norm(sec, nsec int) (int, int) {
	for nsec >= 1000000000 {
		nsec -= 1000000000
		sec++
	}
	for nsec < 0 {
		nsec += 1000000000
		sec--
	}
	return sec, nsec
}

func __go_time_Unix(sec, nsec int) __goTime {
	s, n := __go_time_norm(sec, nsec)
	return __goTime{sec: s, nsec: n, loc: __goLoc{name: "UTC", offset: 0}}
}

func __go_time_UnixMilli(ms int) __goTime {
	return __go_time_Unix(ms/1000, (ms%1000)*1000000)
}

func __go_time_UnixMicro(us int) __goTime {
	return __go_time_Unix(us/1000000, (us%1000000)*1000)
}

func __go_time_Date(year, month, day, hour, minu, second, nsec int, loc __goLoc) __goTime {
	ms := __go_date_utc(year, month-1, day, hour, minu, second)
	sec := ms/1000 - loc.offset
	return __goTime{sec: sec, nsec: nsec, loc: loc}
}

func __go_time_Now() __goTime {
	dt := __go_wall_now()
	return __goTime{sec: dt.seconds, nsec: dt.nanoseconds, loc: __goLoc{name: "UTC", offset: 0}}
}

func __go_time_FixedZone(name string, offset int) __goLoc {
	return __goLoc{name: name, offset: offset}
}

func (t __goTime) __localMs() int {
	return (t.sec + t.loc.offset) * 1000
}
func (t __goTime) Year() int       { return __go_date_year(__go_date_new(t.__localMs())) }
func (t __goTime) Day() int        { return __go_date_day(__go_date_new(t.__localMs())) }
func (t __goTime) Hour() int       { return __go_date_hour(__go_date_new(t.__localMs())) }
func (t __goTime) Minute() int     { return __go_date_min(__go_date_new(t.__localMs())) }
func (t __goTime) Second() int     { return __go_date_sec(__go_date_new(t.__localMs())) }
func (t __goTime) Nanosecond() int { return t.nsec }
func (t __goTime) Unix() int       { return t.sec }
func (t __goTime) UnixNano() int   { return t.sec*1000000000 + t.nsec }
func (t __goTime) UnixMilli() int  { return t.sec*1000 + t.nsec/1000000 }
func (t __goTime) UnixMicro() int  { return t.sec*1000000 + t.nsec/1000 }
func (t __goTime) UTC() __goTime {
	return __goTime{sec: t.sec, nsec: t.nsec, loc: __goLoc{name: "UTC", offset: 0}}
}
func (t __goTime) Location() __goLoc { return t.loc }
func (t __goTime) IsZero() bool      { return t.sec == 0 && t.nsec == 0 }
func (t __goTime) Before(u __goTime) bool {
	return t.sec < u.sec || (t.sec == u.sec && t.nsec < u.nsec)
}
func (t __goTime) After(u __goTime) bool {
	return t.sec > u.sec || (t.sec == u.sec && t.nsec > u.nsec)
}
func (t __goTime) Equal(u __goTime) bool { return t.sec == u.sec && t.nsec == u.nsec }
func (t __goTime) Add(d int) __goTime {
	s, n := __go_time_norm(t.sec+d/1000000000, t.nsec+d%1000000000)
	return __goTime{sec: s, nsec: n, loc: t.loc}
}
func (t __goTime) Sub(u __goTime) int {
	return (t.sec-u.sec)*1000000000 + (t.nsec - u.nsec)
}

func main() {}
"#;

/// Go-source runtime prelude for `net/url`, wrapping the shared WHATWG `web:url`
/// host (`__go_url_*` builtins) into Go's `URL`/`Values`/`Userinfo` shapes.
/// `##` delimiters because the URL fragment code contains `"#"`.
const GO_NETURL_PRELUDE: &str = r##"package main

import "strings"
import "sort"

type __goUser struct {
	name    string
	pass    string
	hasPass bool
}

func (u __goUser) Username() string         { return u.name }
func (u __goUser) Password() (string, bool) { return u.pass, u.hasPass }
func (u __goUser) String() string {
	if u.hasPass {
		return u.name + ":" + u.pass
	}
	return u.name
}

type __goValues map[string][]string

func (v __goValues) Get(k string) string {
	if x, ok := v[k]; ok && len(x) > 0 {
		return x[0]
	}
	return ""
}
func (v __goValues) Set(k, val string) { v[k] = []string{val} }
func (v __goValues) Add(k, val string) { v[k] = append(v[k], val) }
func (v __goValues) Del(k string)      { delete(v, k) }
func (v __goValues) Encode() string {
	keys := []string{}
	for k := range v {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	res := ""
	for _, k := range keys {
		for _, val := range v[k] {
			if len(res) > 0 {
				res += "&"
			}
			res += __go_url_qesc(k) + "=" + __go_url_qesc(val)
		}
	}
	return res
}

type __goURL struct {
	Scheme   string
	Host     string
	Path     string
	RawQuery string
	Fragment string
	User     __goUser
	raw      string
}

func (u __goURL) Query() __goValues { return __go_url_parse_query(u.RawQuery) }
func (u __goURL) String() string {
	s := ""
	if len(u.Scheme) > 0 {
		s += u.Scheme + "://"
		if len(u.User.name) > 0 {
			s += u.User.String() + "@"
		}
		s += u.Host
	}
	s += u.Path
	if len(u.RawQuery) > 0 {
		s += "?" + u.RawQuery
	}
	if len(u.Fragment) > 0 {
		s += "#" + u.Fragment
	}
	return s
}
func (u __goURL) ResolveReference(ref __goURL) __goURL {
	r, _ := __go_url_parse_with_base(ref.raw, u.String())
	return r
}

func __go_url_qesc(s string) string {
	return strings.ReplaceAll(__go_url_esc(s), "%20", "+")
}

func __go_url_parse_query(raw string) __goValues {
	v := __goValues{}
	if len(raw) == 0 {
		return v
	}
	for _, p := range strings.Split(raw, "&") {
		if len(p) == 0 {
			continue
		}
		i := strings.Index(p, "=")
		key := p
		val := ""
		if i >= 0 {
			key = p[:i]
			val = p[i+1:]
		}
		key = __go_url_unesc(strings.ReplaceAll(key, "+", " "))
		val = __go_url_unesc(strings.ReplaceAll(val, "+", " "))
		v[key] = append(v[key], val)
	}
	return v
}

func __go_url_parse_with_base(s, base string) (__goURL, error) {
	o := __go_url_parse(s, base)
	absolute := strings.Contains(s, "://")
	scheme := ""
	host := ""
	if absolute {
		scheme = o.protocol
		if len(scheme) > 0 {
			scheme = scheme[:len(scheme)-1]
		}
		host = o.host
	}
	rawq := o.search
	if len(rawq) > 0 {
		rawq = rawq[1:]
	}
	frag := o.hash
	if len(frag) > 0 {
		frag = frag[1:]
	}
	user := __goUser{name: o.username, pass: o.password, hasPass: len(o.password) > 0}
	return __goURL{Scheme: scheme, Host: host, Path: o.pathname, RawQuery: rawq, Fragment: frag, User: user, raw: s}, nil
}

func __go_url_Parse(s string) (__goURL, error) {
	return __go_url_parse_with_base(s, "http://__vybe_base_/")
}

func __go_url_ParseRequestURI(s string) (__goURL, error) {
	return __go_url_parse_with_base(s, "http://__vybe_base_/")
}

func __go_url_PathEscape(s string) string {
	return strings.ReplaceAll(__go_url_esc(s), "%2F", "/")
}

func __go_url_PathUnescape(s string) (string, error) {
	return __go_url_unesc(s), nil
}

func __go_url_JoinPath(base string, elems []string) string {
	res := base
	for _, e := range elems {
		if len(res) > 0 && res[len(res)-1] != '/' {
			res += "/"
		}
		res += e
	}
	parts := strings.Split(res, "/")
	out := []string{}
	for _, p := range parts {
		if p == ".." && len(out) > 0 && out[len(out)-1] != ".." && out[len(out)-1] != "" {
			out = out[:len(out)-1]
		} else if p != "." {
			out = append(out, p)
		}
	}
	return strings.Join(out, "/")
}

func __go_url_User(name string) __goUser {
	return __goUser{name: name, pass: "", hasPass: false}
}

func __go_url_UserPassword(name, pass string) __goUser {
	return __goUser{name: name, pass: pass, hasPass: true}
}

func main() {}
"##;

/// Go-source runtime prelude for `bytes.Buffer` (string-backed accumulator).
const GO_BYTES_PRELUDE: &str = r#"package main

type __goBuffer struct {
	data string
}

func (b *__goBuffer) WriteString(s string) (int, error) {
	b.data = b.data + s
	return len(s), nil
}
func (b *__goBuffer) Write(p []byte) (int, error) {
	b.data = b.data + string(p)
	return len(p), nil
}
func (b *__goBuffer) WriteByte(c byte) error {
	b.data = b.data + string(rune(c))
	return nil
}
func (b *__goBuffer) String() string { return b.data }
func (b *__goBuffer) Len() int       { return len(b.data) }
func (b *__goBuffer) Reset()         { b.data = "" }
func (b *__goBuffer) Bytes() []byte  { return []byte(b.data) }

func main() {}
"#;

/// Go-source runtime prelude for `log/slog` (structured logging). Handlers write
/// formatted `level`/`msg`/`key=val` lines to their `io.Writer` (a `bytes.Buffer`
/// in the tests). Levels are a named int type so `Level.String()` works.
const GO_SLOG_PRELUDE: &str = r#"package main

import "fmt"

type __goLevel int

func (l __goLevel) String() string {
	if l <= -4 {
		return "DEBUG"
	}
	if l < 4 {
		return "INFO"
	}
	if l < 8 {
		return "WARN"
	}
	return "ERROR"
}

func __go_slog_LevelDebug() __goLevel { return __goLevel(-4) }
func __go_slog_LevelInfo() __goLevel  { return __goLevel(0) }
func __go_slog_LevelWarn() __goLevel  { return __goLevel(4) }
func __go_slog_LevelError() __goLevel { return __goLevel(8) }

type __goAttr struct {
	key string
	val string
}

func __go_slog_Int(k string, v int) __goAttr     { return __goAttr{key: k, val: fmt.Sprintf("%v", v)} }
func __go_slog_Int64(k string, v int64) __goAttr { return __goAttr{key: k, val: fmt.Sprintf("%v", v)} }
func __go_slog_String(k, v string) __goAttr      { return __goAttr{key: k, val: v} }
func __go_slog_Float64(k string, v float64) __goAttr {
	return __goAttr{key: k, val: fmt.Sprintf("%v", v)}
}
func __go_slog_Duration(k string, v int) __goAttr {
	return __goAttr{key: k, val: fmt.Sprintf("%v", v)}
}
func __go_slog_Any(k string, v int) __goAttr { return __goAttr{key: k, val: fmt.Sprintf("%v", v)} }
func __go_slog_Bool(k string, v bool) __goAttr {
	val := "false"
	if v {
		val = "true"
	}
	return __goAttr{key: k, val: val}
}
func __go_slog_Group(k string, attrs []__goAttr) __goAttr {
	s := ""
	for _, a := range attrs {
		if len(s) > 0 {
			s = s + " "
		}
		s = s + a.key + "=" + a.val
	}
	return __goAttr{key: k, val: s}
}

type __goHandlerOptions struct {
	Level     __goLevel
	AddSource bool
}

type __goSlogHandler struct {
	w     *__goBuffer
	level int
}

type __goSlogLogger struct {
	h *__goSlogHandler
}

func __go_slog_optlevel(opts *__goHandlerOptions) int {
	if opts != nil {
		return int(opts.Level)
	}
	return 0
}
func __go_slog_NewTextHandler(w *__goBuffer, opts *__goHandlerOptions) *__goSlogHandler {
	return &__goSlogHandler{w: w, level: __go_slog_optlevel(opts)}
}
func __go_slog_NewJSONHandler(w *__goBuffer, opts *__goHandlerOptions) *__goSlogHandler {
	return &__goSlogHandler{w: w, level: __go_slog_optlevel(opts)}
}
func __go_slog_New(h *__goSlogHandler) *__goSlogLogger {
	return &__goSlogLogger{h: h}
}
func __go_slog_Default() *__goSlogLogger {
	return &__goSlogLogger{h: &__goSlogHandler{w: &__goBuffer{data: ""}, level: 0}}
}

func (l *__goSlogLogger) __emit(level int, name, msg string, attrs []__goAttr) {
	if level < l.h.level {
		return
	}
	line := "level=" + name + " msg=" + msg
	for _, a := range attrs {
		line = line + " " + a.key + "=" + a.val
	}
	line = line + "\n"
	l.h.w.WriteString(line)
}
func (l *__goSlogLogger) Info(msg string, attrs ...__goAttr)  { l.__emit(0, "INFO", msg, attrs) }
func (l *__goSlogLogger) Debug(msg string, attrs ...__goAttr) { l.__emit(-4, "DEBUG", msg, attrs) }
func (l *__goSlogLogger) Warn(msg string, attrs ...__goAttr)  { l.__emit(4, "WARN", msg, attrs) }
func (l *__goSlogLogger) Error(msg string, attrs ...__goAttr) { l.__emit(8, "ERROR", msg, attrs) }
func (l *__goSlogLogger) LogAttrs(ctx any, level __goLevel, msg string, attrs ...__goAttr) {
	l.__emit(int(level), level.String(), msg, attrs)
}

func main() {}
"#;

/// Go-source runtime prelude for the `slices` and `maps` packages — pure
/// generic algorithms over slices/maps (type params erase, so no element
/// coercion). Rewritten from `slices.*` / `maps.*` in the walker.
const GO_SLICES_MAPS_PRELUDE: &str = r#"package main

func __go_slices_Contains[T any](s []T, v T) bool {
	for _, x := range s {
		if x == v {
			return true
		}
	}
	return false
}
func __go_slices_Index[T any](s []T, v T) int {
	for i, x := range s {
		if x == v {
			return i
		}
	}
	return -1
}
func __go_slices_IndexFunc[T any](s []T, f func(T) bool) int {
	for i, x := range s {
		if f(x) {
			return i
		}
	}
	return -1
}
func __go_slices_Equal[T any](a, b []T) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}
func __go_slices_Compare[T any](a, b []T) int {
	n := len(a)
	if len(b) < n {
		n = len(b)
	}
	for i := 0; i < n; i++ {
		if a[i] < b[i] {
			return -1
		}
		if a[i] > b[i] {
			return 1
		}
	}
	if len(a) < len(b) {
		return -1
	}
	if len(a) > len(b) {
		return 1
	}
	return 0
}
func __go_slices_Clone[T any](s []T) []T {
	r := []T{}
	for _, x := range s {
		r = append(r, x)
	}
	return r
}
func __go_slices_Compact[T any](s []T) []T {
	r := []T{}
	for i, x := range s {
		if i == 0 || x != s[i-1] {
			r = append(r, x)
		}
	}
	return r
}
func __go_slices_Delete[T any](s []T, i, j int) []T {
	r := []T{}
	for k, x := range s {
		if k < i || k >= j {
			r = append(r, x)
		}
	}
	return r
}
func __go_slices_Insert[T any](s []T, i int, vals []T) []T {
	r := []T{}
	for k := 0; k < i; k++ {
		r = append(r, s[k])
	}
	for _, v := range vals {
		r = append(r, v)
	}
	for k := i; k < len(s); k++ {
		r = append(r, s[k])
	}
	return r
}
func __go_slices_Replace[T any](s []T, i, j int, vals []T) []T {
	r := []T{}
	for k := 0; k < i; k++ {
		r = append(r, s[k])
	}
	for _, v := range vals {
		r = append(r, v)
	}
	for k := j; k < len(s); k++ {
		r = append(r, s[k])
	}
	return r
}
func __go_slices_Grow[T any](s []T, n int) []T { return s }
func __go_slices_Clip[T any](s []T) []T        { return s }
func __go_slices_BinarySearch[T any](s []T, target T) (int, bool) {
	lo, hi := 0, len(s)
	for lo < hi {
		mid := (lo + hi) >> 1
		if s[mid] < target {
			lo = mid + 1
		} else {
			hi = mid
		}
	}
	found := lo < len(s) && s[lo] == target
	return lo, found
}
func __go_slices_BinarySearchFunc[T any, E any](s []T, target E, cmp func(T, E) int) (int, bool) {
	lo, hi := 0, len(s)
	for lo < hi {
		mid := (lo + hi) >> 1
		if cmp(s[mid], target) < 0 {
			lo = mid + 1
		} else {
			hi = mid
		}
	}
	found := lo < len(s) && cmp(s[lo], target) == 0
	return lo, found
}

func __go_maps_Clone[K any, V any](m map[K]V) map[K]V {
	r := map[K]V{}
	for k, v := range m {
		r[k] = v
	}
	return r
}
func __go_maps_Copy[K any, V any](dst, src map[K]V) {
	for k, v := range src {
		dst[k] = v
	}
}
func __go_maps_DeleteFunc[K any, V any](m map[K]V, f func(K, V) bool) {
	for k, v := range m {
		if f(k, v) {
			delete(m, k)
		}
	}
}

func main() {}
"#;

/// Go-source runtime prelude for `sync/atomic` function-style ops. The VM is a
/// single logical thread, so these are plain pointer reads/writes; every typed
/// variant (Int32/Int64/Uint32/Uint64) maps to the same helper.
const GO_ATOMIC_PRELUDE: &str = r#"package main

func __go_atomic_Load(p *int64) int64 { return *p }
func __go_atomic_Store(p *int64, v int64) { *p = v }
func __go_atomic_Add(p *int64, delta int64) int64 {
	*p += delta
	return *p
}
func __go_atomic_Swap(p *int64, v int64) int64 {
	old := *p;
	*p = v
	return old
}
func __go_atomic_CAS(p *int64, old, repl int64) bool {
	if *p == old {
		*p = repl
		return true
	}
	return false
}

func main() {}
"#;

/// Go-source runtime prelude for the string-based `strconv` helpers (ParseBool,
/// CanBackquote) — pure string logic, no numeric primitives needed.
const GO_STRCONV_PRELUDE: &str = r#"package main

import "strings"

func __go_strconv_ParseBool(s string) (bool, error) {
	if s == "1" || s == "t" || s == "T" || s == "TRUE" || s == "true" || s == "True" {
		return true, nil
	}
	if s == "0" || s == "f" || s == "F" || s == "FALSE" || s == "false" || s == "False" {
		return false, nil
	}
	return false, "invalid syntax"
}

func __go_strconv_CanBackquote(s string) bool {
	if strings.Contains(s, "`") {
		return false
	}
	for _, c := range s {
		if c == '\n' || c == '\r' || c == '\\' {
			return false
		}
		if c < ' ' && c != '\t' {
			return false
		}
	}
	return true
}

func main() {}
"#;

/// Go-source runtime prelude for composite `strings` helpers that compose the
/// already-wired primitives (`Contains`, `Index`, `HasPrefix`, slicing, `range`).
/// Rewritten from `strings.<Name>` in the walker (see `go_rewrite_strings_call`).
const GO_STRINGS_PRELUDE: &str = r#"package main

import "strings"

func __go_strings_TrimPrefix(s, prefix string) string {
	if strings.HasPrefix(s, prefix) {
		return s[len(prefix):]
	}
	return s
}

func __go_strings_TrimSuffix(s, suffix string) string {
	if strings.HasSuffix(s, suffix) {
		return s[:len(s)-len(suffix)]
	}
	return s
}

func __go_strings_CutPrefix(s, prefix string) (string, bool) {
	if strings.HasPrefix(s, prefix) {
		return s[len(prefix):], true
	}
	return s, false
}

func __go_strings_CutSuffix(s, suffix string) (string, bool) {
	if strings.HasSuffix(s, suffix) {
		return s[:len(s)-len(suffix)], true
	}
	return s, false
}

func __go_strings_Cut(s, sep string) (string, string, bool) {
	i := strings.Index(s, sep)
	if i < 0 {
		return s, "", false
	}
	return s[:i], s[i+len(sep):], true
}

func __go_strings_Replace(s, old, repl string, n int) string {
	if n < 0 {
		return strings.ReplaceAll(s, old, repl)
	}
	res := ""
	for n > 0 {
		i := strings.Index(s, old)
		if i < 0 {
			break
		}
		res += s[:i] + repl
		s = s[i+len(old):]
		n--
	}
	return res + s
}

func __go_strings_ContainsRune(s string, r rune) bool {
	return strings.Contains(s, string(r))
}

func __go_strings_ContainsAny(s, chars string) bool {
	for _, c := range s {
		if strings.Contains(chars, string(c)) {
			return true
		}
	}
	return false
}

func __go_strings_ContainsFunc(s string, f func(rune) bool) bool {
	for _, c := range s {
		if f(c) {
			return true
		}
	}
	return false
}

func __go_strings_IndexByte(s string, b byte) int {
	return strings.Index(s, string(rune(b)))
}

func __go_strings_IndexRune(s string, r rune) int {
	return strings.Index(s, string(r))
}

func __go_strings_IndexAny(s, chars string) int {
	for i, c := range s {
		if strings.Contains(chars, string(c)) {
			return i
		}
	}
	return -1
}

func __go_strings_IndexFunc(s string, f func(rune) bool) int {
	for i, c := range s {
		if f(c) {
			return i
		}
	}
	return -1
}

func __go_strings_LastIndexByte(s string, b byte) int {
	return strings.LastIndex(s, string(rune(b)))
}

func __go_strings_LastIndexAny(s, chars string) int {
	res := -1
	for i, c := range s {
		if strings.Contains(chars, string(c)) {
			res = i
		}
	}
	return res
}

func __go_strings_LastIndexFunc(s string, f func(rune) bool) int {
	res := -1
	for i, c := range s {
		if f(c) {
			res = i
		}
	}
	return res
}

func __go_strings_TrimLeft(s, cutset string) string {
	for len(s) > 0 && strings.Contains(cutset, string(rune(s[0]))) {
		s = s[1:]
	}
	return s
}

func __go_strings_TrimRight(s, cutset string) string {
	for len(s) > 0 && strings.Contains(cutset, string(rune(s[len(s)-1]))) {
		s = s[:len(s)-1]
	}
	return s
}

func __go_strings_TrimCutset(s, cutset string) string {
	return __go_strings_TrimRight(__go_strings_TrimLeft(s, cutset), cutset)
}

func __go_strings_EqualFold(s, t string) bool {
	return strings.ToLower(s) == strings.ToLower(t)
}

func __go_strings_Map(f func(rune) rune, s string) string {
	res := ""
	for _, c := range s {
		m := f(c)
		if m >= 0 {
			res += string(m)
		}
	}
	return res
}

func __go_strings_Fields(s string) []string {
	res := []string{}
	cur := ""
	for _, c := range s {
		if c == ' ' || c == '\t' || c == '\n' || c == '\r' {
			if len(cur) > 0 {
				res = append(res, cur)
				cur = ""
			}
		} else {
			cur += string(c)
		}
	}
	if len(cur) > 0 {
		res = append(res, cur)
	}
	return res
}

func __go_strings_FieldsFunc(s string, f func(rune) bool) []string {
	res := []string{}
	cur := ""
	for _, c := range s {
		if f(c) {
			if len(cur) > 0 {
				res = append(res, cur)
				cur = ""
			}
		} else {
			cur += string(c)
		}
	}
	if len(cur) > 0 {
		res = append(res, cur)
	}
	return res
}

func __go_strings_SplitN(s, sep string, n int) []string {
	if n == 0 {
		return []string{}
	}
	if n < 0 {
		return strings.Split(s, sep)
	}
	res := []string{}
	for n > 1 {
		i := strings.Index(s, sep)
		if i < 0 {
			break
		}
		res = append(res, s[:i])
		s = s[i+len(sep):]
		n--
	}
	res = append(res, s)
	return res
}

func __go_strings_SplitAfter(s, sep string) []string {
	res := []string{}
	for len(sep) > 0 {
		i := strings.Index(s, sep)
		if i < 0 {
			break
		}
		res = append(res, s[:i+len(sep)])
		s = s[i+len(sep):]
	}
	res = append(res, s)
	return res
}

func __go_strings_SplitAfterN(s, sep string, n int) []string {
	if n == 0 {
		return []string{}
	}
	if n < 0 {
		return __go_strings_SplitAfter(s, sep)
	}
	res := []string{}
	for n > 1 {
		i := strings.Index(s, sep)
		if i < 0 {
			break
		}
		res = append(res, s[:i+len(sep)])
		s = s[i+len(sep):]
		n--
	}
	res = append(res, s)
	return res
}

func main() {}
"#;

/// Walk a Go source string into its raw (pre-normalization) parts.
fn walk_go_source(source: &str) -> Result<(String, Vec<Statement>, Vec<Import>), String> {
    let pairs =
        GoParser::parse(Rule::program, source).map_err(|e| format!("Go parse error: {}", e))?;

    let mut body = Vec::new();
    let mut imports = Vec::new();
    let mut package_name = String::new();

    for top in pairs {
        if top.as_rule() == Rule::EOI {
            continue;
        }
        let inner = match top.as_rule() {
            Rule::program => top.into_inner(),
            _ => {
                if let Some(stmt) = walk_top_level(top)? {
                    body.push(stmt);
                }
                continue;
            }
        };
        for pair in inner {
            match pair.as_rule() {
                Rule::EOI => continue,
                Rule::package_clause => {
                    package_name = walk_package_clause(pair)?;
                }
                Rule::import_declarations => {
                    for imp in pair.into_inner() {
                        if imp.as_rule() == Rule::import_declaration {
                            imports.push(walk_import(imp)?);
                        }
                    }
                }
                _ => {
                    if let Some(stmt) = walk_top_level(pair)? {
                        body.push(stmt);
                    }
                }
            }
        }
    }

    Ok((package_name, body, imports))
}

/// Go-source runtime prelude for the `errors` package + `fmt.Errorf`.
///
/// Errors are modeled as a value struct `__goError{message, wrap, errs}` with
/// `Error()`/`Unwrap()` methods — a plain Go value, so distinct literals stay
/// `!=` (matching Go's pointer-based `errors.New` distinctness under the VM's
/// object-identity `==`) while value type assertions still resolve. The
/// package functions (`errors.New/Is/Unwrap/Join`, `fmt.Errorf`) are rewritten
/// in the walker to call these helpers; `errors.As` is rewritten with
/// type-assertion closures at the call site (it is generic over the target
/// type).
const GO_ERRORS_PRELUDE: &str = r#"package main

type __goError struct {
	message string
	wrap    error
	errs    []error
}

func (e __goError) Error() string { return e.message }
func (e __goError) Unwrap() error { return e.wrap }

func __go_new_error(message string, wrap error, errs []error) error {
	return __goError{message: message, wrap: wrap, errs: errs}
}

func __go_errors_unwrap(err error) error {
	if err == nil {
		return nil
	}
	if ge, ok := err.(__goError); ok {
		return ge.wrap
	}
	return nil
}

func __go_errors_is(err error, target error) bool {
	if target == nil {
		return false
	}
	worklist := []error{err}
	for len(worklist) > 0 {
		cur := worklist[len(worklist)-1]
		worklist = worklist[:len(worklist)-1]
		if cur == nil {
			continue
		}
		if cur == target {
			return true
		}
		if ge, ok := cur.(__goError); ok {
			if ge.errs != nil {
				worklist = append(worklist, ge.errs...)
			}
			if ge.wrap != nil {
				worklist = append(worklist, ge.wrap)
			}
		}
	}
	return false
}

func __go_errors_join(errs []error) error {
	filtered := []error{}
	for _, e := range errs {
		if e != nil {
			filtered = append(filtered, e)
		}
	}
	if len(filtered) == 0 {
		return nil
	}
	msg := ""
	for i, e := range filtered {
		if i > 0 {
			msg = msg + "\n"
		}
		msg = msg + e.Error()
	}
	return __go_new_error(msg, nil, filtered)
}

func __go_errors_as(err error, match func(error) bool, assign func(error)) bool {
	worklist := []error{err}
	for len(worklist) > 0 {
		cur := worklist[len(worklist)-1]
		worklist = worklist[:len(worklist)-1]
		if cur == nil {
			continue
		}
		if match(cur) {
			assign(cur)
			return true
		}
		if ge, ok := cur.(__goError); ok {
			if ge.errs != nil {
				worklist = append(worklist, ge.errs...)
			}
			if ge.wrap != nil {
				worklist = append(worklist, ge.wrap)
			}
		}
	}
	return false
}
"#;

#[derive(Clone, Default)]
struct GoFunctionSignature {
    params: Vec<Option<String>>,
    return_type: Option<String>,
}

#[derive(Clone, Default)]
struct GoNormalizeEnv {
    value_types: HashMap<String, String>,
    fixed_arrays: HashMap<String, String>,
    slice_caps: HashMap<String, Expression>,
    slice_views: HashMap<String, GoSliceViewInfo>,
    struct_infos: HashMap<String, GoStructInfo>,
    named_types: HashMap<String, String>,
    type_names: HashSet<String>,
    return_type: Option<String>,
    panic_value_name: Option<String>,
    has_panic_name: Option<String>,
    in_defer_name: Option<String>,
    recover_fn_name: Option<String>,
    owns_panic_state: bool,
}

#[derive(Clone)]
struct GoSliceViewInfo {
    base: Expression,
    start: Expression,
    end: Option<Expression>,
}

#[derive(Clone, Default)]
struct GoStructInfo {
    field_order: Vec<String>,
    member_names: HashSet<String>,
    method_names: HashSet<String>,
    member_types: HashMap<String, String>,
    embedded_fields: Vec<(String, String)>,
}

#[derive(Default)]
struct GoNormalizeState {
    next_temp: usize,
}

struct GoSignatureInfo {
    params: Vec<Param>,
    return_type: Option<String>,
    named_results: Vec<Param>,
}

fn normalize_go_module(mut module: Module) -> Module {
    module.body = merge_go_struct_decls(&module.body);
    let signatures = collect_go_function_signatures(&module.body);
    let globals = collect_go_global_fixed_arrays(&module.body, &signatures);
    let struct_infos = collect_go_struct_infos(&module.body);
    let named_types = collect_go_named_types(&module.body);
    let type_names = collect_go_type_names(&module.body);
    let mut state = GoNormalizeState::default();
    let mut env = GoNormalizeEnv {
        value_types: HashMap::new(),
        fixed_arrays: globals.clone(),
        slice_caps: HashMap::new(),
        slice_views: HashMap::new(),
        struct_infos,
        named_types,
        type_names,
        return_type: None,
        panic_value_name: None,
        has_panic_name: None,
        in_defer_name: None,
        recover_fn_name: None,
        owns_panic_state: false,
    };

    let mut normalized = Vec::with_capacity(module.body.len());
    for stmt in &module.body {
        normalized.extend(normalize_go_statement(
            stmt,
            &mut env,
            &signatures,
            &mut state,
        ));
    }
    module.body = go_lower_module_init_functions(normalized, &mut state);
    module
}

fn go_lower_module_init_functions(
    body: Vec<Statement>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let mut lowered = Vec::with_capacity(body.len());
    let mut init_calls = Vec::new();

    for stmt in body {
        match stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                body,
                modifiers,
                handles,
                is_async,
                is_generator,
                is_sub,
            } if name == "init" => {
                let hidden_name = fresh_go_temp(state, "__go_init");
                lowered.push(Statement::new(StmtKind::FunctionDecl {
                    name: hidden_name.clone(),
                    params,
                    return_type,
                    body,
                    modifiers,
                    handles,
                    is_async,
                    is_generator,
                    is_sub,
                }));
                init_calls.push(Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident(&hidden_name)),
                        args: Vec::new(),
                        optional: false,
                    },
                ))));
            }
            _ => lowered.push(stmt),
        }
    }

    lowered.extend(init_calls);
    lowered
}

fn collect_go_function_signatures(body: &[Statement]) -> HashMap<String, GoFunctionSignature> {
    let mut signatures = HashMap::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::FunctionDecl {
                name,
                params,
                return_type,
                ..
            } => {
                signatures.insert(
                    name.clone(),
                    GoFunctionSignature {
                        params: params.iter().map(|param| param.type_hint.clone()).collect(),
                        return_type: return_type.clone(),
                    },
                );
            }
            StmtKind::StructDecl { members, .. } => {
                for member in members {
                    if let ClassMember::Method(stmt) = member {
                        if let StmtKind::FunctionDecl {
                            name,
                            params,
                            return_type,
                            ..
                        } = &stmt.kind
                        {
                            signatures.insert(
                                name.clone(),
                                GoFunctionSignature {
                                    params: params
                                        .iter()
                                        .map(|param| param.type_hint.clone())
                                        .collect(),
                                    return_type: return_type.clone(),
                                },
                            );
                        }
                    }
                }
            }
            _ => {}
        }
    }
    signatures
}

fn collect_go_type_names(body: &[Statement]) -> HashSet<String> {
    let mut type_names = HashSet::new();
    for stmt in body {
        match &stmt.kind {
            StmtKind::StructDecl { name, .. }
            | StmtKind::InterfaceDecl { name, .. }
            | StmtKind::EnumDecl { name, .. }
            | StmtKind::ClassDecl { name, .. } => {
                type_names.insert(name.clone());
            }
            _ => {
                if let Some((name, _)) = go_extract_named_type_marker(stmt) {
                    type_names.insert(name);
                }
            }
        }
    }
    type_names
}

fn collect_go_named_types(body: &[Statement]) -> HashMap<String, String> {
    let mut named_types = HashMap::new();
    for stmt in body {
        if let Some((name, underlying)) = go_extract_named_type_marker(stmt) {
            named_types.insert(name, underlying);
        }
    }
    named_types
}

fn collect_go_struct_infos(body: &[Statement]) -> HashMap<String, GoStructInfo> {
    let mut infos = HashMap::new();
    for stmt in body {
        let StmtKind::StructDecl { name, members, .. } = &stmt.kind else {
            continue;
        };
        let info = infos
            .entry(name.clone())
            .or_insert_with(GoStructInfo::default);
        for member in members {
            match member {
                ClassMember::Field {
                    name, type_hint, ..
                } => {
                    info.field_order.push(name.clone());
                    info.member_names.insert(name.clone());
                    if let Some(type_name) = type_hint.clone() {
                        info.member_types.insert(name.clone(), type_name.clone());
                        if go_embedded_field_name(&type_name).as_deref() == Some(name.as_str()) {
                            info.embedded_fields.push((name.clone(), type_name));
                        }
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name, return_type, ..
                    } = &stmt.kind
                    {
                        info.member_names.insert(name.clone());
                        info.method_names.insert(name.clone());
                        if let Some(type_name) = return_type.clone() {
                            info.member_types.insert(name.clone(), type_name);
                        }
                    }
                }
                _ => {}
            }
        }
    }
    infos
}

fn merge_go_struct_decls(body: &[Statement]) -> Vec<Statement> {
    let mut first_index: HashMap<String, usize> = HashMap::new();
    for (index, stmt) in body.iter().enumerate() {
        if let StmtKind::StructDecl { name, .. } = &stmt.kind {
            first_index.entry(name.clone()).or_insert(index);
        }
    }

    let mut emitted = std::collections::HashSet::new();
    let mut merged_body = Vec::with_capacity(body.len());

    for (index, stmt) in body.iter().enumerate() {
        match &stmt.kind {
            StmtKind::StructDecl { name, .. } => {
                if first_index.get(name) != Some(&index) || !emitted.insert(name.clone()) {
                    continue;
                }

                let mut merged = stmt.clone();
                if let StmtKind::StructDecl {
                    interfaces,
                    members,
                    ..
                } = &mut merged.kind
                {
                    for later in body.iter().skip(index + 1) {
                        if let StmtKind::StructDecl {
                            name: later_name,
                            interfaces: later_interfaces,
                            members: later_members,
                            ..
                        } = &later.kind
                        {
                            if later_name == name {
                                members.extend(later_members.clone());
                                for interface in later_interfaces {
                                    if !interfaces.iter().any(|existing| existing == interface) {
                                        interfaces.push(interface.clone());
                                    }
                                }
                            }
                        }
                    }
                }

                merged_body.push(merged);
            }
            _ => merged_body.push(stmt.clone()),
        }
    }

    merged_body
}

fn collect_go_global_fixed_arrays(
    body: &[Statement],
    signatures: &HashMap<String, GoFunctionSignature>,
) -> HashMap<String, String> {
    let env = GoNormalizeEnv::default();
    let mut globals = HashMap::new();

    for stmt in body {
        if let StmtKind::VarDecl { declarations, .. } = &stmt.kind {
            for decl in declarations {
                if let Some((name, type_name)) = go_decl_fixed_array_binding(decl, &env, signatures)
                {
                    globals.insert(name, type_name);
                }
            }
        }
    }

    globals
}

fn normalize_go_block(
    stmts: &[Statement],
    base_env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let mut env = base_env.clone();
    let mut normalized = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        normalized.extend(normalize_go_statement(stmt, &mut env, signatures, state));
    }
    normalized
}

fn normalize_go_function_body(
    stmts: &[Statement],
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    if env.recover_fn_name.is_none() {
        env.panic_value_name = Some(fresh_go_temp(state, "__go_panic_value"));
        env.has_panic_name = Some(fresh_go_temp(state, "__go_has_panic"));
        env.in_defer_name = Some(fresh_go_temp(state, "__go_in_defer"));
        env.recover_fn_name = Some(fresh_go_temp(state, "__go_recover"));
        env.owns_panic_state = true;
    } else {
        env.owns_panic_state = false;
    }

    let mut named_result: Option<Param> = None;
    let mut body_stmts = Vec::with_capacity(stmts.len());
    for stmt in stmts {
        if named_result.is_none() {
            if let Some(param) = go_extract_named_result_marker(stmt) {
                env.value_types.insert(
                    param.name.clone(),
                    param
                        .type_hint
                        .clone()
                        .unwrap_or_else(|| "object".to_string()),
                );
                named_result = Some(param);
                continue;
            }
        }
        body_stmts.push(stmt.clone());
    }

    let mut normalized = Vec::with_capacity(stmts.len());
    for stmt in &body_stmts {
        normalized.extend(normalize_go_statement(stmt, env, signatures, state));
    }

    let (normalized, final_return) = if let Some(param) = named_result {
        go_lower_named_result_body(normalized, &param, state)
    } else {
        (normalized, None)
    };

    lower_go_defer_body(normalized, env, signatures, state, final_return)
}

fn lower_go_defer_body(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    final_return: Option<Expression>,
) -> Vec<Statement> {
    let panic_value_name = env
        .panic_value_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_panic_value"));
    let has_panic_name = env
        .has_panic_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_has_panic"));
    let in_defer_name = env
        .in_defer_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_in_defer"));
    let recover_fn_name = env
        .recover_fn_name
        .clone()
        .unwrap_or_else(|| fresh_go_temp(state, "__go_recover"));
    let stack_name = fresh_go_temp(state, "__go_defer_stack");
    let (lowered_body, has_defer) =
        lower_go_defer_statements(body, env, signatures, state, &stack_name, false);

    let panic_value_decl = go_defer_temp_decl(panic_value_name.clone(), None, Expression::null());
    let has_panic_decl = go_defer_temp_decl(has_panic_name.clone(), None, Expression::bool(false));
    let in_defer_decl = go_defer_temp_decl(in_defer_name.clone(), None, Expression::bool(false));
    let recover_fn_decl = go_defer_temp_decl(
        recover_fn_name,
        None,
        Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(Expression::ident(&in_defer_name)),
                    right: Box::new(Expression::ident(&has_panic_name)),
                }),
                then_body: vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&has_panic_name)],
                        value: Expression::bool(false),
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(&panic_value_name)))),
                ],
                elifs: Vec::new(),
                else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                    Expression::null(),
                )))]),
            })]),
            is_async: false,
            captures: Vec::new(),
        }),
    );

    let panic_state_decls = if env.owns_panic_state {
        vec![
            panic_value_decl,
            has_panic_decl,
            in_defer_decl,
            recover_fn_decl,
        ]
    } else {
        Vec::new()
    };

    if !has_defer {
        let mut body = panic_state_decls;
        body.extend(lowered_body);
        if let Some(expr) = final_return {
            body.push(Statement::new(StmtKind::Return(Some(expr))));
        }
        return body;
    }

    let stack_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(stack_name.clone()),
            type_hint: None,
            init: Some(Expression::null()),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let drain_name = fresh_go_temp(state, "__go_defer_fn");
    let drain_loop = Statement::new(StmtKind::While {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::ident(&stack_name)),
            right: Box::new(Expression::null()),
        }),
        body: vec![
            go_defer_temp_decl(
                drain_name.clone(),
                None,
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "fn".to_string(),
                    null_safe: false,
                }),
            ),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&stack_name)],
                value: Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "next".to_string(),
                    null_safe: false,
                }),
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&in_defer_name)],
                value: Expression::bool(true),
            }),
            Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(&drain_name)),
                args: Vec::new(),
                optional: false,
            }))),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&in_defer_name)],
                value: Expression::bool(false),
            }),
        ],
        else_body: None,
    });

    let panic_catch_name = fresh_go_temp(state, "__go_panic_exc");

    let mut body = panic_state_decls;
    body.extend([
        stack_decl,
        Statement::new(StmtKind::Try {
            body: lowered_body,
            catches: vec![CatchClause {
                types: Vec::new(),
                var_name: Some(panic_catch_name.clone()),
                stack_var: None,
                body: vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&panic_value_name)],
                        value: Expression::ident(&panic_catch_name),
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&has_panic_name)],
                        value: Expression::bool(true),
                    }),
                ],
                when_clause: None,
            }],
            else_body: None,
            finally: Some(vec![drain_loop]),
        }),
        Statement::new(StmtKind::If {
            cond: Expression::ident(&has_panic_name),
            then_body: vec![Statement::new(StmtKind::Throw {
                expr: Some(Expression::ident(&panic_value_name)),
                cause: None,
            })],
            elifs: Vec::new(),
            else_body: None,
        }),
    ]);
    if let Some(expr) = final_return {
        body.push(Statement::new(StmtKind::Return(Some(expr))));
    }
    body
}

fn go_extract_named_result_marker(stmt: &Statement) -> Option<Param> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_named_result")
        || args.len() != 2
    {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let type_hint = go_type_name_from_expr(&args[1].value)?;
    Some(Param {
        name: name.clone(),
        type_hint: Some(type_hint),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    })
}

fn go_extract_named_type_marker(stmt: &Statement) -> Option<(String, String)> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_named_type")
        || args.len() != 2
    {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let type_name = go_type_name_from_expr(&args[1].value)?;
    Some((name.clone(), type_name))
}

fn go_lower_named_result_body(
    body: Vec<Statement>,
    result: &Param,
    state: &mut GoNormalizeState,
) -> (Vec<Statement>, Option<Expression>) {
    // Use a sentinel string only as a rewrite marker for go_rewrite_named_result_returns.
    // We no longer throw/catch this sentinel at runtime — instead we use a
    // `while(true) { ...; break }` loop so that `return X` inside any branch
    // compiles to `result = X; break`, which emits a clean `BR(N)` that
    // correctly unwinds the label stack.  Throwing a sentinel inside an IF
    // body left extra BLOCK labels on the label_stack that THROW does not
    // restore, corrupting the outer catch handler lookup.
    let sentinel = fresh_go_temp(state, "__go_named_return");
    let result_name = result.name.clone();
    let result_type = result
        .type_hint
        .clone()
        .unwrap_or_else(|| "object".to_string());

    let rewritten_body = go_rewrite_named_result_returns(body, &result_name, &sentinel);
    let result_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(result_name.clone()),
            type_hint: Some(result_type.clone()),
            init: Some(go_zero_value_expr(&result_type)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    // Wrap the rewritten body in `while(true) { ...; break }`.
    // Every `return X` in rewritten_body was turned into `result=X; break`.
    // At end of function body (bare `return`) we also emit an implicit break.
    // Real panics (user-thrown exceptions) propagate directly to the outer
    // defer try/catch because no try/catch is interposed here.
    let mut while_body = rewritten_body;
    // Ensure the while always exits: append an implicit break so fall-through
    // at end of body exits the loop rather than looping forever.
    while_body.push(Statement::new(StmtKind::Break(BreakTarget::Implicit)));

    (
        vec![
            result_decl,
            Statement::new(StmtKind::While {
                cond: Expression::bool(true),
                body: while_body,
                else_body: None,
            }),
        ],
        Some(Expression::ident(&result_name)),
    )
}

fn go_rewrite_named_result_returns(
    body: Vec<Statement>,
    result_name: &str,
    sentinel: &str,
) -> Vec<Statement> {
    let mut rewritten = Vec::with_capacity(body.len());
    for stmt in body {
        rewritten.extend(go_rewrite_named_result_return_stmt(
            stmt,
            result_name,
            sentinel,
        ));
    }
    rewritten
}

fn go_rewrite_named_result_return_stmt(
    stmt: Statement,
    result_name: &str,
    sentinel: &str,
) -> Vec<Statement> {
    match stmt.kind {
        StmtKind::Return(expr) => {
            let mut rewritten = Vec::new();
            if let Some(expr) = expr {
                rewritten.push(Statement::new(StmtKind::Assign {
                    targets: vec![Expression::ident(result_name)],
                    value: expr,
                }));
            }
            // Break out of the enclosing while(true) loop cleanly.
            // This compiles to BR(N) which correctly unwinds the label stack
            // regardless of how many nested BLOCKs (e.g. from if-statements)
            // are active at the break site.
            rewritten.push(Statement::new(StmtKind::Break(BreakTarget::Implicit)));
            rewritten
        }
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(
            go_rewrite_named_result_returns(body, result_name, sentinel),
        ))],
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => vec![Statement::new(StmtKind::If {
            cond,
            then_body: go_rewrite_named_result_returns(then_body, result_name, sentinel),
            elifs: elifs
                .into_iter()
                .map(|(cond, body)| {
                    (
                        cond,
                        go_rewrite_named_result_returns(body, result_name, sentinel),
                    )
                })
                .collect(),
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => vec![Statement::new(StmtKind::For {
            init,
            cond,
            update,
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
        })],
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => vec![Statement::new(StmtKind::ForIn {
            var,
            key,
            iter,
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            of,
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
            is_async,
        })],
        StmtKind::While {
            cond,
            body,
            else_body,
        } => vec![Statement::new(StmtKind::While {
            cond,
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        StmtKind::DoWhile { body, cond, until } => vec![Statement::new(StmtKind::DoWhile {
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            cond,
            until,
        })],
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => vec![Statement::new(StmtKind::Switch {
            expr,
            cases: cases
                .into_iter()
                .map(|case| SwitchCase {
                    conditions: case.conditions,
                    body: go_rewrite_named_result_returns(case.body, result_name, sentinel),
                })
                .collect(),
            default: default
                .map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => vec![Statement::new(StmtKind::Try {
            body: go_rewrite_named_result_returns(body, result_name, sentinel),
            catches: catches
                .into_iter()
                .map(|catch| CatchClause {
                    types: catch.types,
                    var_name: catch.var_name,
                    stack_var: catch.stack_var,
                    body: go_rewrite_named_result_returns(catch.body, result_name, sentinel),
                    when_clause: catch.when_clause,
                })
                .collect(),
            else_body: else_body
                .map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
            finally: finally
                .map(|body| go_rewrite_named_result_returns(body, result_name, sentinel)),
        })],
        _ => vec![stmt],
    }
}

fn lower_go_defer_statements(
    body: Vec<Statement>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    in_loop: bool,
) -> (Vec<Statement>, bool) {
    let mut lowered = Vec::with_capacity(body.len());
    let mut has_defer = false;
    let mut loop_local_names = HashSet::new();
    let empty_loop_local_names = HashSet::new();

    for stmt in body {
        if let Some(expr) = go_extract_defer_expr(&stmt) {
            let frozen_names = if in_loop {
                &loop_local_names
            } else {
                &empty_loop_local_names
            };
            lowered.extend(go_lower_defer_stmt(
                expr,
                env,
                signatures,
                state,
                stack_name,
                frozen_names,
                in_loop,
            ));
            has_defer = true;
            continue;
        }

        let (next_stmt, nested_has_defer) =
            lower_go_defer_statement(stmt, env, signatures, state, stack_name, in_loop);
        if in_loop {
            go_collect_block_declared_names(&next_stmt, &mut loop_local_names);
        }
        lowered.push(next_stmt);
        has_defer |= nested_has_defer;
    }

    (lowered, has_defer)
}

fn lower_go_defer_statement(
    stmt: Statement,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    in_loop: bool,
) -> (Statement, bool) {
    match stmt.kind {
        StmtKind::Block(body) => {
            let (body, has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
            (Statement::new(StmtKind::Block(body)), has_defer)
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let (next_then, mut has_defer) =
                lower_go_defer_statements(then_body, env, signatures, state, stack_name, in_loop);
            let mut next_elifs = Vec::with_capacity(elifs.len());
            for (elif_cond, elif_body) in elifs {
                let (next_body, nested_has_defer) = lower_go_defer_statements(
                    elif_body, env, signatures, state, stack_name, in_loop,
                );
                next_elifs.push((elif_cond, next_body));
                has_defer |= nested_has_defer;
            }
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::If {
                    cond,
                    then_body: next_then,
                    elifs: next_elifs,
                    else_body: next_else,
                }),
                has_defer,
            )
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let (body, has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, true);
            (
                Statement::new(StmtKind::For {
                    init,
                    cond,
                    update,
                    body,
                }),
                has_defer,
            )
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => {
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, true);
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::ForIn {
                    var,
                    key,
                    iter,
                    body,
                    of,
                    else_body: next_else,
                    is_async,
                }),
                has_defer,
            )
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, true);
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::While {
                    cond,
                    body,
                    else_body: next_else,
                }),
                has_defer,
            )
        }
        StmtKind::DoWhile { body, cond, until } => {
            let (body, has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, true);
            (
                Statement::new(StmtKind::DoWhile { body, cond, until }),
                has_defer,
            )
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            let mut has_defer = false;
            let next_cases = cases
                .into_iter()
                .map(|case| {
                    let (body, nested_has_defer) = lower_go_defer_statements(
                        case.body, env, signatures, state, stack_name, in_loop,
                    );
                    has_defer |= nested_has_defer;
                    SwitchCase {
                        conditions: case.conditions,
                        body,
                    }
                })
                .collect();
            let next_default = if let Some(body) = default {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Switch {
                    expr,
                    cases: next_cases,
                    default: next_default,
                }),
                has_defer,
            )
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
            let next_catches = catches
                .into_iter()
                .map(|catch| {
                    let (body, nested_has_defer) = lower_go_defer_statements(
                        catch.body, env, signatures, state, stack_name, in_loop,
                    );
                    has_defer |= nested_has_defer;
                    CatchClause {
                        types: catch.types,
                        var_name: catch.var_name,
                        stack_var: catch.stack_var,
                        body,
                        when_clause: catch.when_clause,
                    }
                })
                .collect();
            let next_else = if let Some(body) = else_body {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            let next_finally = if let Some(body) = finally {
                let (body, nested_has_defer) =
                    lower_go_defer_statements(body, env, signatures, state, stack_name, in_loop);
                has_defer |= nested_has_defer;
                Some(body)
            } else {
                None
            };
            (
                Statement::new(StmtKind::Try {
                    body,
                    catches: next_catches,
                    else_body: next_else,
                    finally: next_finally,
                }),
                has_defer,
            )
        }
        _ => (stmt, false),
    }
}

fn go_extract_defer_expr(stmt: &Statement) -> Option<Expression> {
    let StmtKind::Expr(expr) = &stmt.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_defer") || args.len() != 1
    {
        return None;
    }
    Some(args[0].value.clone())
}

fn go_lower_defer_stmt(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
    stack_name: &str,
    frozen_names: &HashSet<String>,
    in_loop: bool,
) -> Vec<Statement> {
    let mut stmts = Vec::new();
    let mut loop_snapshot_captures: Vec<String> = Vec::new();

    let deferred_expr = match expr.kind {
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            let deferred_callee = match callee.as_ref() {
                Expression {
                    kind:
                        ExprKind::Member {
                            object,
                            field,
                            null_safe,
                        },
                    ..
                } => {
                    let receiver_type = go_expr_type_hint(object, env, signatures);
                    let deferred_object = if matches!(object.as_ref().kind, ExprKind::Ident(_))
                        && receiver_type.is_none()
                    {
                        object.as_ref().clone()
                    } else {
                        let temp_name = fresh_go_temp(state, "__go_defer_recv");
                        stmts.push(go_defer_temp_decl(
                            temp_name.clone(),
                            receiver_type,
                            object.as_ref().clone(),
                        ));
                        if in_loop {
                            loop_snapshot_captures.push(temp_name.clone());
                        }
                        Expression::ident(&temp_name)
                    };
                    Expression::new(ExprKind::Member {
                        object: Box::new(deferred_object),
                        field: field.clone(),
                        null_safe: *null_safe,
                    })
                }
                _ => {
                    let deferred_value =
                        go_freeze_defer_lambda_captures(callee.as_ref().clone(), frozen_names);
                    let temp_name = fresh_go_temp(state, "__go_defer_fn");
                    stmts.push(go_defer_temp_decl(
                        temp_name.clone(),
                        go_expr_type_hint(&deferred_value, env, signatures),
                        deferred_value,
                    ));
                    if in_loop {
                        loop_snapshot_captures.push(temp_name.clone());
                    }
                    Expression::ident(&temp_name)
                }
            };

            let deferred_args = args
                .into_iter()
                .map(|arg| {
                    let temp_name = fresh_go_temp(state, "__go_defer_arg");
                    let value = go_wrap_fixed_array_copy(arg.value, env, signatures);
                    stmts.push(go_defer_temp_decl(
                        temp_name.clone(),
                        go_expr_type_hint(&value, env, signatures),
                        value,
                    ));
                    if in_loop {
                        loop_snapshot_captures.push(temp_name.clone());
                    }
                    Argument {
                        value: Expression::ident(&temp_name),
                        name: arg.name,
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    }
                })
                .collect();

            Expression::new(ExprKind::Call {
                callee: Box::new(deferred_callee),
                args: deferred_args,
                optional,
            })
        }
        _ => expr,
    };

    // Build the zero-arg closure that will be stored on the defer stack.
    //
    // Non-loop case: use empty explicit captures so the compiler routes through
    // the outer function's shared env (parent_shared_env_slot path).  The
    // __go_defer_arg* / __go_defer_fn* temps are set-once per function call,
    // so reading them from the shared env at drain time always gives the
    // value they had at defer-registration time.
    //
    // Loop case: the same shared-env slot is OVERWRITTEN on every iteration,
    // so all closures would see the last iteration's value.  We fix this by
    // wrapping the inner lambda in an IIFE that takes the loop-iteration
    // variables as parameters (shadowing the outer slots).  Each IIFE call
    // gets its own activation record, so the inner lambda closes over a
    // freshly-snapshotted copy of the values for that iteration.
    let inner_lambda = Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Expr(deferred_expr))]),
        is_async: false,
        captures: Vec::new(),
    });

    let closure = if in_loop && !loop_snapshot_captures.is_empty() {
        let wrapper_params: Vec<Param> = loop_snapshot_captures
            .iter()
            .map(|name| Param {
                name: name.clone(),
                type_hint: None,
                default: None,
                is_rest: false,
                pass_by: PassBy::Value,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect();
        let wrapper = Expression::new(ExprKind::Lambda {
            params: wrapper_params,
            body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(inner_lambda)))]),
            is_async: false,
            captures: Vec::new(),
        });
        let iife_args: Vec<Argument> = loop_snapshot_captures
            .iter()
            .map(|name| Argument {
                value: Expression::ident(name),
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect();
        Expression::new(ExprKind::Call {
            callee: Box::new(wrapper),
            args: iife_args,
            optional: false,
        })
    } else {
        inner_lambda
    };
    stmts.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(stack_name)],
        value: Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("fn"),
                value: closure,
            },
            ObjectProperty::KeyValue {
                key: Expression::string("next"),
                value: Expression::ident(stack_name),
            },
        ])),
    }));
    stmts
}

fn go_collect_block_declared_names(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                go_collect_binding_pattern_names(&decl.pattern, names);
            }
        }
        StmtKind::FunctionDecl { name, .. } => {
            names.insert(name.clone());
        }
        _ => {}
    }
}

fn go_collect_binding_pattern_names(pattern: &BindingPattern, names: &mut HashSet<String>) {
    match pattern {
        BindingPattern::Ident(name) => {
            names.insert(name.clone());
        }
        BindingPattern::Array(elements) => {
            for element in elements {
                if let ArrayPatternElem::Pattern(pattern, _) = element {
                    go_collect_binding_pattern_names(pattern, names);
                }
            }
        }
        BindingPattern::Object(properties) => {
            for property in properties {
                if let Some(pattern) = &property.value {
                    go_collect_binding_pattern_names(pattern, names);
                }
            }
        }
    }
}

fn go_rewrite_immediate_lambda_ref_captures(
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Lambda {
        params,
        body,
        is_async,
        captures,
    } = &callee.kind
    else {
        return None;
    };

    let param_names: HashSet<String> = params.iter().map(|param| param.name.clone()).collect();
    let mut local_names = HashSet::new();
    go_collect_lambda_declared_names(body, &mut local_names);

    let mut ref_names = HashSet::new();
    go_collect_lambda_ref_idents(body, &mut ref_names);

    let mut captured_ref_names = ref_names
        .into_iter()
        .filter(|name| !param_names.contains(name) && !local_names.contains(name))
        .collect::<Vec<_>>();
    captured_ref_names.sort();

    if captured_ref_names.is_empty() {
        return None;
    }

    let mut replacements = HashMap::new();
    let mut next_params = params.clone();
    let mut next_args = args.to_vec();

    for name in captured_ref_names {
        let temp_name = fresh_go_temp(state, "__go_ref_capture");
        let pointee_type = go_expr_type_hint(&Expression::ident(&name), env, signatures);
        next_params.push(Param {
            name: temp_name.clone(),
            type_hint: pointee_type.map(|type_name| format!("*{}", type_name.trim())),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        });
        next_args.push(Argument::positional(Expression::new(ExprKind::RefOf(
            Box::new(PlaceExpr::Ident(name.clone())),
        ))));
        replacements.insert(name, temp_name);
    }

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: next_params,
            body: go_rewrite_lambda_ref_body(body, &replacements),
            is_async: *is_async,
            captures: captures.clone(),
        })),
        args: next_args,
        optional,
    }))
}

fn go_collect_lambda_declared_names(body: &LambdaBody, names: &mut HashSet<String>) {
    if let LambdaBody::Block(stmts) = body {
        for stmt in stmts {
            go_collect_stmt_declared_names_recursive(stmt, names);
        }
    }
}

fn go_collect_stmt_declared_names_recursive(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                go_collect_binding_pattern_names(&decl.pattern, names);
            }
        }
        StmtKind::FunctionDecl { name, .. } => {
            names.insert(name.clone());
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
        }
        StmtKind::If {
            then_body,
            elifs,
            else_body,
            ..
        } => {
            for stmt in then_body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
            for (_, body) in elifs {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
        }
        StmtKind::For { init, body, .. } => {
            if let Some(init) = init {
                go_collect_stmt_declared_names_recursive(init, names);
            }
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
        }
        StmtKind::ForIn {
            var,
            key,
            body,
            else_body,
            ..
        } => {
            names.insert(var.clone());
            if let Some(key) = key {
                names.insert(key.clone());
            }
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
        }
        StmtKind::While {
            body, else_body, ..
        } => {
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_declared_names_recursive(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, .. } => {
            for stmt in body {
                go_collect_stmt_declared_names_recursive(stmt, names);
            }
        }
        _ => {}
    }
}

fn go_collect_lambda_ref_idents(body: &LambdaBody, names: &mut HashSet<String>) {
    match body {
        LambdaBody::Expr(expr) => go_collect_expr_ref_idents(expr, names),
        LambdaBody::Block(stmts) => {
            for stmt in stmts {
                go_collect_stmt_ref_idents(stmt, names);
            }
        }
    }
}

fn go_collect_stmt_ref_idents(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(expr) => go_collect_expr_ref_idents(expr, names),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                go_collect_expr_ref_idents(expr, names);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                go_collect_expr_ref_idents(expr, names);
            }
            if let Some(cause) = cause {
                go_collect_expr_ref_idents(cause, names);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                go_collect_expr_ref_idents(target, names);
            }
            go_collect_expr_ref_idents(value, names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_collect_expr_ref_idents(target, names);
            go_collect_expr_ref_idents(value, names);
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    go_collect_expr_ref_idents(init, names);
                }
            }
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_collect_expr_ref_idents(cond, names);
            for stmt in then_body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            for (cond, body) in elifs {
                go_collect_expr_ref_idents(cond, names);
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                go_collect_stmt_ref_idents(init, names);
            }
            if let Some(cond) = cond {
                go_collect_expr_ref_idents(cond, names);
            }
            if let Some(update) = update {
                go_collect_expr_ref_idents(update, names);
            }
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            go_collect_expr_ref_idents(iter, names);
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_collect_expr_ref_idents(cond, names);
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            go_collect_expr_ref_idents(cond, names);
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                go_collect_stmt_ref_idents(stmt, names);
            }
            for catch in catches {
                if let Some(expr) = &catch.when_clause {
                    go_collect_expr_ref_idents(expr, names);
                }
                for stmt in &catch.body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    go_collect_stmt_ref_idents(stmt, names);
                }
            }
        }
        _ => {}
    }
}

fn go_collect_expr_ref_idents(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::RefOf(place) => {
            if let PlaceExpr::Ident(name) = place.as_ref() {
                names.insert(name.clone());
            }
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            if let ExprKind::Ident(name) = &expr.kind {
                names.insert(name.clone());
            }
            go_collect_expr_ref_idents(expr, names);
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) | ExprKind::Cast { expr, .. } => {
            go_collect_expr_ref_idents(expr, names)
        }
        ExprKind::Binary { left, right, .. } => {
            go_collect_expr_ref_idents(left, names);
            go_collect_expr_ref_idents(right, names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_collect_expr_ref_idents(cond, names);
            go_collect_expr_ref_idents(then, names);
            go_collect_expr_ref_idents(else_, names);
        }
        ExprKind::Member { object, .. } => go_collect_expr_ref_idents(object, names),
        ExprKind::Index { object, index, .. } => {
            go_collect_expr_ref_idents(object, names);
            go_collect_expr_ref_idents(index, names);
        }
        ExprKind::Assign { target, value } => {
            go_collect_expr_ref_idents(target, names);
            go_collect_expr_ref_idents(value, names);
        }
        ExprKind::Call { callee, args, .. } => {
            go_collect_expr_ref_idents(callee, names);
            for arg in args {
                go_collect_expr_ref_idents(&arg.value, names);
            }
        }
        ExprKind::Lambda { body, .. } => go_collect_lambda_ref_idents(body, names),
        _ => {}
    }
}

fn go_rewrite_lambda_ref_body(
    body: &LambdaBody,
    replacements: &HashMap<String, String>,
) -> LambdaBody {
    match body {
        LambdaBody::Expr(expr) => {
            LambdaBody::Expr(Box::new(go_rewrite_expr_ref_idents(expr, replacements)))
        }
        LambdaBody::Block(stmts) => LambdaBody::Block(
            stmts
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
        ),
    }
}

fn go_rewrite_stmt_ref_idents(
    stmt: &Statement,
    replacements: &HashMap<String, String>,
) -> Statement {
    match &stmt.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Expr(go_rewrite_expr_ref_idents(
            expr,
            replacements,
        ))),
        StmtKind::Return(expr) => Statement::new(StmtKind::Return(
            expr.as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
        )),
        StmtKind::Throw { expr, cause } => Statement::new(StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
            cause: cause
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
        }),
        StmtKind::Assign { targets, value } => Statement::new(StmtKind::Assign {
            targets: targets
                .iter()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements))
                .collect(),
            value: go_rewrite_expr_ref_idents(value, replacements),
        }),
        StmtKind::CompoundAssign { target, op, value } => {
            Statement::new(StmtKind::CompoundAssign {
                target: go_rewrite_expr_ref_idents(target, replacements),
                op: *op,
                value: go_rewrite_expr_ref_idents(value, replacements),
            })
        }
        StmtKind::VarDecl { declarations, kind } => Statement::new(StmtKind::VarDecl {
            declarations: declarations
                .iter()
                .map(|decl| VarDeclarator {
                    pattern: decl.pattern.clone(),
                    type_hint: decl.type_hint.clone(),
                    init: decl
                        .init
                        .as_ref()
                        .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
                    array_bounds: decl.array_bounds.clone(),
                    with_events: decl.with_events,
                })
                .collect(),
            kind: kind.clone(),
        }),
        StmtKind::Block(body) => Statement::new(StmtKind::Block(
            body.iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
        )),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => Statement::new(StmtKind::If {
            cond: go_rewrite_expr_ref_idents(cond, replacements),
            then_body: then_body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            elifs: elifs
                .iter()
                .map(|(cond, body)| {
                    (
                        go_rewrite_expr_ref_idents(cond, replacements),
                        body.iter()
                            .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                            .collect(),
                    )
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
        }),
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => Statement::new(StmtKind::For {
            init: init
                .as_ref()
                .map(|stmt| Box::new(go_rewrite_stmt_ref_idents(stmt, replacements))),
            cond: cond
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
            update: update
                .as_ref()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
        }),
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => Statement::new(StmtKind::ForIn {
            var: var.clone(),
            key: key.clone(),
            iter: go_rewrite_expr_ref_idents(iter, replacements),
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            of: *of,
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
            is_async: *is_async,
        }),
        StmtKind::While {
            cond,
            body,
            else_body,
        } => Statement::new(StmtKind::While {
            cond: go_rewrite_expr_ref_idents(cond, replacements),
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
        }),
        StmtKind::DoWhile { body, cond, until } => Statement::new(StmtKind::DoWhile {
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            cond: go_rewrite_expr_ref_idents(cond, replacements),
            until: *until,
        }),
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => Statement::new(StmtKind::Try {
            body: body
                .iter()
                .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                .collect(),
            catches: catches
                .iter()
                .map(|catch| CatchClause {
                    types: catch.types.clone(),
                    var_name: catch.var_name.clone(),
                    stack_var: catch.stack_var.clone(),
                    body: catch
                        .body
                        .iter()
                        .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                        .collect(),
                    when_clause: catch
                        .when_clause
                        .as_ref()
                        .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
                })
                .collect(),
            else_body: else_body.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
            finally: finally.as_ref().map(|body| {
                body.iter()
                    .map(|stmt| go_rewrite_stmt_ref_idents(stmt, replacements))
                    .collect()
            }),
        }),
        _ => stmt.clone(),
    }
}

fn go_rewrite_expr_ref_idents(
    expr: &Expression,
    replacements: &HashMap<String, String>,
) -> Expression {
    match &expr.kind {
        ExprKind::RefOf(place) => {
            if let PlaceExpr::Ident(name) = place.as_ref() {
                if let Some(replacement) = replacements.get(name) {
                    return Expression::ident(replacement);
                }
            }
            expr.clone()
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: inner,
        } => {
            if let ExprKind::Ident(name) = &inner.kind {
                if let Some(replacement) = replacements.get(name) {
                    return Expression::ident(replacement);
                }
            }
            Expression::new(ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr: Box::new(go_rewrite_expr_ref_idents(inner, replacements)),
            })
        }
        ExprKind::Unary { op, expr: inner } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(go_rewrite_expr_ref_idents(inner, replacements)),
        }),
        ExprKind::RefLoad(inner) => Expression::new(ExprKind::RefLoad(Box::new(
            go_rewrite_expr_ref_idents(inner, replacements),
        ))),
        ExprKind::Cast {
            expr: inner,
            type_name,
        } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_expr_ref_idents(inner, replacements)),
            type_name: type_name.clone(),
        }),
        ExprKind::Binary { left, op, right } => Expression::new(ExprKind::Binary {
            left: Box::new(go_rewrite_expr_ref_idents(left, replacements)),
            op: *op,
            right: Box::new(go_rewrite_expr_ref_idents(right, replacements)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_expr_ref_idents(cond, replacements)),
            then: Box::new(go_rewrite_expr_ref_idents(then, replacements)),
            else_: Box::new(go_rewrite_expr_ref_idents(else_, replacements)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_expr_ref_idents(object, replacements)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_expr_ref_idents(object, replacements)),
            index: Box::new(go_rewrite_expr_ref_idents(index, replacements)),
            null_safe: *null_safe,
        }),
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(go_rewrite_expr_ref_idents(target, replacements)),
            value: Box::new(go_rewrite_expr_ref_idents(value, replacements)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_expr_ref_idents(callee, replacements)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: go_rewrite_expr_ref_idents(&arg.value, replacements),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|expr| go_rewrite_expr_ref_idents(expr, replacements)),
                    value: go_rewrite_expr_ref_idents(&element.value, replacements),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        )),
        ExprKind::Object(properties) => Expression::new(ExprKind::Object(
            properties
                .iter()
                .map(|property| match property {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: go_rewrite_expr_ref_idents(key, replacements),
                        value: go_rewrite_expr_ref_idents(value, replacements),
                    },
                    ObjectProperty::Spread(value) => {
                        ObjectProperty::Spread(go_rewrite_expr_ref_idents(value, replacements))
                    }
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: go_rewrite_expr_ref_idents(key, replacements),
                        value: go_rewrite_expr_ref_idents(value, replacements),
                    },
                    _ => property.clone(),
                })
                .collect(),
        )),
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .iter()
                .map(|value| go_rewrite_expr_ref_idents(value, replacements))
                .collect(),
        )),
        ExprKind::Sequence(values) => Expression::new(ExprKind::Sequence(
            values
                .iter()
                .map(|value| go_rewrite_expr_ref_idents(value, replacements))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => Expression::new(ExprKind::Lambda {
            params: params.clone(),
            body: go_rewrite_lambda_ref_body(body, replacements),
            is_async: *is_async,
            captures: captures.clone(),
        }),
        _ => expr.clone(),
    }
}

fn go_freeze_defer_lambda_captures(expr: Expression, frozen_names: &HashSet<String>) -> Expression {
    let ExprKind::Lambda {
        params,
        body,
        is_async,
        mut captures,
    } = expr.kind
    else {
        return expr;
    };

    if !frozen_names.is_empty() {
        let mut used_names = HashSet::new();
        go_collect_lambda_body_idents(&body, &mut used_names);
        let param_names: HashSet<String> = params.iter().map(|param| param.name.clone()).collect();
        let mut frozen_capture_names = used_names
            .into_iter()
            .filter(|name| frozen_names.contains(name) && !param_names.contains(name))
            .collect::<Vec<_>>();
        frozen_capture_names.sort();
        for name in frozen_capture_names {
            if !captures.iter().any(|capture| capture == &name) {
                captures.push(name);
            }
        }
    }

    Expression::new(ExprKind::Lambda {
        params,
        body,
        is_async,
        captures,
    })
}

fn go_collect_lambda_body_idents(body: &LambdaBody, names: &mut HashSet<String>) {
    match body {
        LambdaBody::Expr(expr) => go_collect_expr_idents(expr, names),
        LambdaBody::Block(stmts) => {
            for stmt in stmts {
                go_collect_stmt_idents(stmt, names);
            }
        }
    }
}

fn go_collect_stmt_idents(stmt: &Statement, names: &mut HashSet<String>) {
    match &stmt.kind {
        StmtKind::Expr(expr) => go_collect_expr_idents(expr, names),
        StmtKind::Return(expr) => {
            if let Some(expr) = expr {
                go_collect_expr_idents(expr, names);
            }
        }
        StmtKind::Throw { expr, cause } => {
            if let Some(expr) = expr {
                go_collect_expr_idents(expr, names);
            }
            if let Some(cause) = cause {
                go_collect_expr_idents(cause, names);
            }
        }
        StmtKind::Assign { targets, value } => {
            for target in targets {
                go_collect_expr_idents(target, names);
            }
            go_collect_expr_idents(value, names);
        }
        StmtKind::CompoundAssign { target, value, .. } => {
            go_collect_expr_idents(target, names);
            go_collect_expr_idents(value, names);
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            go_collect_expr_idents(cond, names);
            for stmt in then_body {
                go_collect_stmt_idents(stmt, names);
            }
            for (cond, body) in elifs {
                go_collect_expr_idents(cond, names);
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            if let Some(init) = init {
                go_collect_stmt_idents(init, names);
            }
            if let Some(cond) = cond {
                go_collect_expr_idents(cond, names);
            }
            if let Some(update) = update {
                go_collect_expr_idents(update, names);
            }
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
        }
        StmtKind::ForIn {
            iter,
            body,
            else_body,
            ..
        } => {
            go_collect_expr_idents(iter, names);
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => {
            go_collect_expr_idents(cond, names);
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::DoWhile { body, cond, .. } => {
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            go_collect_expr_idents(cond, names);
        }
        StmtKind::Block(body) => {
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            go_collect_expr_idents(expr, names);
            for case in cases {
                for condition in &case.conditions {
                    match condition {
                        CaseCondition::Value(expr) => go_collect_expr_idents(expr, names),
                        CaseCondition::Range { from, to } => {
                            go_collect_expr_idents(from, names);
                            go_collect_expr_idents(to, names);
                        }
                        CaseCondition::Comparison { expr, .. } => {
                            go_collect_expr_idents(expr, names)
                        }
                    }
                }
                for stmt in &case.body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = default {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => {
            for stmt in body {
                go_collect_stmt_idents(stmt, names);
            }
            for catch in catches {
                if let Some(cond) = &catch.when_clause {
                    go_collect_expr_idents(cond, names);
                }
                for stmt in &catch.body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = else_body {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
            if let Some(body) = finally {
                for stmt in body {
                    go_collect_stmt_idents(stmt, names);
                }
            }
        }
        StmtKind::VarDecl { declarations, .. } => {
            for decl in declarations {
                if let Some(init) = &decl.init {
                    go_collect_expr_idents(init, names);
                }
            }
        }
        _ => {}
    }
}

fn go_collect_expr_idents(expr: &Expression, names: &mut HashSet<String>) {
    match &expr.kind {
        ExprKind::Ident(name) => {
            names.insert(name.clone());
        }
        ExprKind::Unary { expr, .. } | ExprKind::RefLoad(expr) | ExprKind::Cast { expr, .. } => {
            go_collect_expr_idents(expr, names)
        }
        ExprKind::AddressOf(name) => {
            names.insert(name.clone());
        }
        ExprKind::Binary { left, right, .. } => {
            go_collect_expr_idents(left, names);
            go_collect_expr_idents(right, names);
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_collect_expr_idents(cond, names);
            go_collect_expr_idents(then, names);
            go_collect_expr_idents(else_, names);
        }
        ExprKind::Member { object, .. } => go_collect_expr_idents(object, names),
        ExprKind::Index { object, index, .. } => {
            go_collect_expr_idents(object, names);
            go_collect_expr_idents(index, names);
        }
        ExprKind::Assign { target, value } => {
            go_collect_expr_idents(target, names);
            go_collect_expr_idents(value, names);
        }
        ExprKind::Call { callee, args, .. } => {
            go_collect_expr_idents(callee, names);
            for arg in args {
                go_collect_expr_idents(&arg.value, names);
            }
        }
        ExprKind::Array(elements) => {
            for element in elements {
                if let Some(key) = &element.key {
                    go_collect_expr_idents(key, names);
                }
                go_collect_expr_idents(&element.value, names);
            }
        }
        ExprKind::Object(properties) => {
            for property in properties {
                match property {
                    ObjectProperty::KeyValue { key, value } => {
                        go_collect_expr_idents(key, names);
                        go_collect_expr_idents(value, names);
                    }
                    ObjectProperty::Spread(value) => go_collect_expr_idents(value, names),
                    ObjectProperty::Computed { key, value } => {
                        go_collect_expr_idents(key, names);
                        go_collect_expr_idents(value, names);
                    }
                    _ => {}
                }
            }
        }
        ExprKind::Tuple(values) | ExprKind::Sequence(values) => {
            for value in values {
                go_collect_expr_idents(value, names);
            }
        }
        ExprKind::Lambda { body, .. } => go_collect_lambda_body_idents(body, names),
        _ => {}
    }
}

fn go_defer_temp_decl(name: String, type_hint: Option<String>, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name),
            type_hint,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn normalize_go_statement(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    if env.recover_fn_name.is_none() {
        env.panic_value_name = Some(fresh_go_temp(state, "__go_panic_value"));
        env.has_panic_name = Some(fresh_go_temp(state, "__go_has_panic"));
        env.in_defer_name = Some(fresh_go_temp(state, "__go_in_defer"));
        env.recover_fn_name = Some(fresh_go_temp(state, "__go_recover"));
    }

    match &stmt.kind {
        StmtKind::FunctionDecl {
            name,
            params,
            return_type,
            body,
            modifiers,
            handles,
            is_async,
            is_generator,
            is_sub,
        } => {
            let mut fn_env = GoNormalizeEnv {
                value_types: env.value_types.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                slice_caps: env.slice_caps.clone(),
                slice_views: env.slice_views.clone(),
                struct_infos: env.struct_infos.clone(),
                named_types: env.named_types.clone(),
                type_names: env.type_names.clone(),
                return_type: return_type.clone(),
                panic_value_name: None,
                has_panic_name: None,
                in_defer_name: None,
                recover_fn_name: None,
                owns_panic_state: false,
            };
            for param in params {
                if let Some(type_hint) = param.type_hint.as_ref() {
                    fn_env
                        .value_types
                        .insert(param.name.clone(), type_hint.clone());
                }
                if let Some(type_hint) = param
                    .type_hint
                    .as_deref()
                    .filter(|hint| go_is_fixed_array_type(hint))
                {
                    fn_env
                        .fixed_arrays
                        .insert(param.name.clone(), type_hint.to_string());
                }
            }

            vec![Statement::new(StmtKind::FunctionDecl {
                name: name.clone(),
                params: params.clone(),
                return_type: return_type.clone(),
                body: normalize_go_function_body(body, &mut fn_env, signatures, state),
                modifiers: modifiers.clone(),
                handles: handles.clone(),
                is_async: *is_async,
                is_generator: *is_generator,
                is_sub: *is_sub,
            })]
        }
        StmtKind::VarDecl { declarations, kind } => {
            let mut normalized = Vec::with_capacity(declarations.len());
            for decl in declarations {
                let mut next_decl = decl.clone();
                if let Some(pattern) = go_single_named_binding_pattern(&next_decl.pattern) {
                    next_decl.pattern = pattern;
                }
                next_decl.init = decl.init.as_ref().map(|expr| {
                    if go_is_two_value_binding_pattern(&decl.pattern) {
                        if let Some(tuple_expr) =
                            go_normalize_map_lookup_tuple_expr(expr, env, signatures, state)
                        {
                            return tuple_expr;
                        }
                    }
                    normalize_go_expr(expr, env, signatures, state)
                });
                next_decl.array_bounds = decl.array_bounds.as_ref().map(|bounds| {
                    bounds
                        .iter()
                        .map(|expr| normalize_go_expr(expr, env, signatures, state))
                        .collect()
                });

                if next_decl.init.is_none()
                    && next_decl
                        .type_hint
                        .as_deref()
                        .is_some_and(go_is_fixed_array_type)
                {
                    next_decl.init = next_decl.type_hint.as_deref().map(go_zero_value_expr);
                } else if next_decl.init.is_none() {
                    if let Some(type_name) = next_decl.type_hint.as_deref() {
                        if let Some(underlying) = env.named_types.get(type_name) {
                            next_decl.init = Some(Expression::new(ExprKind::Cast {
                                expr: Box::new(go_zero_value_expr(underlying)),
                                type_name: type_name.to_string(),
                            }));
                        }
                    }
                } else if let Some(init_expr) = next_decl.init.take() {
                    next_decl.init = Some(go_wrap_fixed_array_copy(init_expr, env, signatures));
                }

                if let Some((name, type_name)) =
                    go_decl_fixed_array_binding(&next_decl, env, signatures)
                {
                    env.fixed_arrays.insert(name, type_name);
                }
                if let Some((name, type_name)) = go_decl_binding_type(&next_decl, env, signatures) {
                    env.value_types.insert(name, type_name);
                }
                if let Some(name) = go_binding_name(&next_decl.pattern) {
                    if let Some(view) = next_decl
                        .init
                        .as_ref()
                        .and_then(|init| go_expr_slice_view(init, env))
                    {
                        if go_slice_view_is_self_referential(&view, &name) {
                            env.slice_views.remove(&name);
                        } else {
                            env.slice_views.insert(name.clone(), view);
                        }
                    } else {
                        env.slice_views.remove(&name);
                    }
                    if let Some(cap_expr) = decl
                        .init
                        .as_ref()
                        .and_then(|init| go_make_slice_capacity_expr(init, env, signatures, state))
                        .or_else(|| {
                            next_decl
                                .init
                                .as_ref()
                                .and_then(|init| go_bound_slice_capacity_expr(init, env))
                        })
                    {
                        env.slice_caps.insert(name, cap_expr);
                    }
                }
                normalized.push(next_decl);
            }

            vec![Statement::new(StmtKind::VarDecl {
                declarations: normalized,
                kind: kind.clone(),
            })]
        }
        StmtKind::Expr(_) if go_extract_named_type_marker(stmt).is_some() => vec![stmt.clone()],
        StmtKind::Expr(expr) => {
            if let Some(panic_expr) = go_extract_panic_expr(expr) {
                vec![Statement::new(StmtKind::Throw {
                    expr: Some(normalize_go_expr(panic_expr, env, signatures, state)),
                    cause: None,
                })]
            } else {
                vec![Statement::new(StmtKind::Expr(normalize_go_expr(
                    expr, env, signatures, state,
                )))]
            }
        }
        StmtKind::Assign { targets, value } => {
            let mut next_value = normalize_go_expr(value, env, signatures, state);
            next_value = go_wrap_fixed_array_copy(next_value, env, signatures);
            if let [target] = targets.as_slice() {
                if let ExprKind::Ident(name) = &target.kind {
                    if let Some(type_name) = go_expr_type_hint(&next_value, env, signatures) {
                        env.value_types.insert(name.clone(), type_name);
                    }
                    if let Some(view) = go_expr_slice_view(&next_value, env) {
                        if go_slice_view_is_self_referential(&view, name) {
                            env.slice_views.remove(name);
                        } else {
                            env.slice_views.insert(name.clone(), view);
                        }
                    } else {
                        env.slice_views.remove(name);
                    }
                    if let Some(cap_expr) =
                        go_make_slice_capacity_expr(value, env, signatures, state)
                            .or_else(|| go_bound_slice_capacity_expr(&next_value, env))
                    {
                        env.slice_caps.insert(name.clone(), cap_expr);
                    }
                }
            }
            vec![Statement::new(StmtKind::Assign {
                targets: targets
                    .iter()
                    .map(|target| normalize_go_lvalue_expr(target, env, signatures, state))
                    .collect(),
                value: next_value,
            })]
        }
        StmtKind::CompoundAssign { target, op, value } => {
            vec![Statement::new(StmtKind::CompoundAssign {
                target: normalize_go_lvalue_expr(target, env, signatures, state),
                op: *op,
                value: normalize_go_expr(value, env, signatures, state),
            })]
        }
        StmtKind::Return(expr) => {
            let next_expr = expr.as_ref().map(|value| {
                let normalized = normalize_go_expr(value, env, signatures, state);
                if env
                    .return_type
                    .as_deref()
                    .is_some_and(go_is_fixed_array_type)
                {
                    go_wrap_fixed_array_copy(normalized, env, signatures)
                } else {
                    normalized
                }
            });
            vec![Statement::new(StmtKind::Return(next_expr))]
        }
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => {
            let next_elifs = elifs
                .iter()
                .map(|(elif_cond, elif_body)| {
                    (
                        normalize_go_expr(elif_cond, env, signatures, state),
                        normalize_go_block(elif_body, env, signatures, state),
                    )
                })
                .collect();
            let next_else = else_body
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state));
            vec![Statement::new(StmtKind::If {
                cond: normalize_go_expr(cond, env, signatures, state),
                then_body: normalize_go_block(then_body, env, signatures, state),
                elifs: next_elifs,
                else_body: next_else,
            })]
        }
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => {
            let next_cases = cases
                .iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(value) => CaseCondition::Value(normalize_go_expr(
                                value, env, signatures, state,
                            )),
                            _ => condition.clone(),
                        })
                        .collect(),
                    body: normalize_go_block(&case.body, env, signatures, state),
                })
                .collect();
            vec![Statement::new(StmtKind::Switch {
                expr: normalize_go_expr(expr, env, signatures, state),
                cases: next_cases,
                default: default
                    .as_ref()
                    .map(|body| normalize_go_block(body, env, signatures, state)),
            })]
        }
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => {
            let mut loop_env = env.clone();
            let next_init = init.as_ref().map(|stmt| {
                Box::new(normalize_go_single_statement(
                    stmt,
                    &mut loop_env,
                    signatures,
                    state,
                ))
            });
            let next_cond = cond
                .as_ref()
                .map(|expr| normalize_go_expr(expr, &loop_env, signatures, state));
            let next_update = update
                .as_ref()
                .map(|expr| normalize_go_expr(expr, &loop_env, signatures, state));
            let next_body = normalize_go_block(body, &loop_env, signatures, state);
            vec![Statement::new(StmtKind::For {
                init: next_init,
                cond: next_cond,
                update: next_update,
                body: next_body,
            })]
        }
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            of,
            else_body,
            is_async,
        } => {
            let next_iter = normalize_go_expr(iter, env, signatures, state);
            if *of && go_expr_type_hint(&next_iter, env, signatures).as_deref() == Some("string") {
                lower_go_string_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else if *of && go_expr_is_fixed_array(&next_iter, env, signatures) {
                lower_go_fixed_array_range(
                    var,
                    key.as_deref(),
                    next_iter,
                    body,
                    env,
                    signatures,
                    state,
                )
            } else {
                vec![Statement::new(StmtKind::ForIn {
                    var: var.clone(),
                    key: key.clone(),
                    iter: next_iter,
                    body: normalize_go_block(body, env, signatures, state),
                    of: *of,
                    else_body: else_body
                        .as_ref()
                        .map(|body| normalize_go_block(body, env, signatures, state)),
                    is_async: *is_async,
                })]
            }
        }
        StmtKind::While {
            cond,
            body,
            else_body,
        } => vec![Statement::new(StmtKind::While {
            cond: normalize_go_expr(cond, env, signatures, state),
            body: normalize_go_block(body, env, signatures, state),
            else_body: else_body
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state)),
        })],
        StmtKind::DoWhile { body, cond, until } => vec![Statement::new(StmtKind::DoWhile {
            body: normalize_go_block(body, env, signatures, state),
            cond: normalize_go_expr(cond, env, signatures, state),
            until: *until,
        })],
        StmtKind::Block(body) => vec![Statement::new(StmtKind::Block(normalize_go_block(
            body, env, signatures, state,
        )))],
        StmtKind::Throw { expr, cause } => vec![Statement::new(StmtKind::Throw {
            expr: expr
                .as_ref()
                .map(|value| normalize_go_expr(value, env, signatures, state)),
            cause: cause
                .as_ref()
                .map(|value| normalize_go_expr(value, env, signatures, state)),
        })],
        StmtKind::StructDecl {
            name,
            interfaces,
            members,
            visibility,
            decorators,
        } => {
            let normalized_members = members
                .iter()
                .map(|member| match member {
                    ClassMember::Method(stmt) => {
                        let normalized_method =
                            normalize_go_single_statement(stmt, env, signatures, state);
                        ClassMember::Method(Box::new(normalized_method))
                    }
                    ClassMember::Field {
                        name,
                        type_hint,
                        init,
                        modifiers,
                        with_events,
                        array_bounds,
                    } => ClassMember::Field {
                        name: name.clone(),
                        type_hint: type_hint.clone(),
                        init: init
                            .as_ref()
                            .map(|expr| normalize_go_expr(expr, env, signatures, state)),
                        modifiers: modifiers.clone(),
                        with_events: *with_events,
                        array_bounds: array_bounds.as_ref().map(|bounds| {
                            bounds
                                .iter()
                                .map(|expr| normalize_go_expr(expr, env, signatures, state))
                                .collect()
                        }),
                    },
                    _ => member.clone(),
                })
                .collect();
            vec![Statement::new(StmtKind::StructDecl {
                name: name.clone(),
                interfaces: interfaces.clone(),
                members: normalized_members,
                visibility: *visibility,
                decorators: decorators.clone(),
            })]
        }
        _ => vec![stmt.clone()],
    }
}

fn normalize_go_single_statement(
    stmt: &Statement,
    env: &mut GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Statement {
    let mut normalized = normalize_go_statement(stmt, env, signatures, state);
    if normalized.len() == 1 {
        normalized.pop().unwrap()
    } else {
        Statement::new(StmtKind::Block(normalized))
    }
}

fn normalize_go_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) => env
            .slice_views
            .get(name)
            .cloned()
            .map(go_materialize_slice_view)
            .unwrap_or_else(|| expr.clone()),
        ExprKind::Binary { op, left, right } => {
            let next_left = normalize_go_expr(left, env, signatures, state);
            let next_right = normalize_go_expr(right, env, signatures, state);
            let normalized_op = if *op == BinOp::Div
                && go_expr_type_hint(&next_left, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_integer_type)
                && go_expr_type_hint(&next_right, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_integer_type)
            {
                BinOp::IDiv
            } else {
                *op
            };

            if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && go_expr_is_fixed_array(&next_left, env, signatures)
                && go_expr_is_fixed_array(&next_right, env, signatures)
            {
                let equal = go_builtin_call("__go_fixed_array_equal", vec![next_left, next_right]);
                if normalized_op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                }
            } else {
                Expression::new(ExprKind::Binary {
                    op: normalized_op,
                    left: Box::new(next_left),
                    right: Box::new(next_right),
                })
            }
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            let next_expr = normalize_go_expr(expr, env, signatures, state);
            if let Some(place) = go_expr_to_place(&next_expr) {
                Expression::new(ExprKind::RefOf(Box::new(place)))
            } else {
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr: Box::new(next_expr),
                })
            }
        }
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        } => Expression::new(ExprKind::RefLoad(Box::new(normalize_go_expr(
            expr, env, signatures, state,
        )))),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(normalize_go_expr(expr, env, signatures, state)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(normalize_go_expr(cond, env, signatures, state)),
            then: Box::new(normalize_go_expr(then, env, signatures, state)),
            else_: Box::new(normalize_go_expr(else_, env, signatures, state)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => {
            if matches!(&object.kind, ExprKind::Ident(name) if name == "time") {
                if let Some(rewritten) = go_rewrite_time_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "slog") {
                if let Some(rewritten) = go_rewrite_slog_member(field) {
                    return rewritten;
                }
            }
            let next_object = normalize_go_expr(object, env, signatures, state);
            let rewritten =
                go_rewrite_promoted_member_access(next_object, field, *null_safe, env, signatures);
            rewritten.unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(normalize_go_expr(object, env, signatures, state)),
                    field: field.clone(),
                    null_safe: *null_safe,
                })
            })
        }
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let next_object = normalize_go_expr(object, env, signatures, state);
            let next_index = normalize_go_expr(index, env, signatures, state);
            if let Some(rewritten) =
                go_rewrite_slice_view_index(&next_object, next_index.clone(), env)
            {
                return rewritten;
            }
            if go_expr_type_hint(&next_object, env, signatures).as_deref() == Some("string") {
                go_member_call(next_object, "charCodeAt", vec![next_index])
            } else if let Some(value_type) = go_expr_type_hint(&next_object, env, signatures)
                .and_then(|type_name| go_map_value_type(&type_name))
            {
                go_build_map_read_expr(next_object, next_index, &value_type)
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(next_object),
                    index: Box::new(next_index),
                    null_safe: *null_safe,
                })
            }
        }
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_lvalue_expr(target, env, signatures, state)),
            value: Box::new(normalize_go_expr(value, env, signatures, state)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            let next_callee = normalize_go_expr(callee, env, signatures, state);
            let signature = match &next_callee.kind {
                ExprKind::Ident(name) => signatures.get(name),
                _ => None,
            };
            let mut next_args = args
                .iter()
                .enumerate()
                .map(|(idx, arg)| {
                    let mut value = normalize_go_expr(&arg.value, env, signatures, state);
                    if signature
                        .and_then(|sig| sig.params.get(idx))
                        .and_then(|hint| hint.as_deref())
                        .is_some_and(go_is_fixed_array_type)
                    {
                        value = go_wrap_fixed_array_copy(value, env, signatures);
                    }
                    Argument {
                        value,
                        name: arg.name.clone(),
                        by_ref: arg.by_ref,
                        spread: arg.spread,
                    }
                })
                .collect::<Vec<_>>();

            if let Some(rewritten_iife) = go_rewrite_immediate_lambda_ref_captures(
                &next_callee,
                &next_args,
                *optional,
                env,
                signatures,
                state,
            ) {
                return rewritten_iife;
            }

            if let Some(rewritten_call) = go_rewrite_named_type_method_call(
                &next_callee,
                &next_args,
                *optional,
                env,
                signatures,
            ) {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_callable_field_member_call(
                &next_callee,
                &next_args,
                *optional,
                env,
                signatures,
            ) {
                return rewritten_call;
            }

            let call_name = go_expr_call_name(&next_callee);

            if let Some(name) = call_name.as_deref() {
                if name == "errors.As" {
                    return go_rewrite_errors_as(&next_args, env, signatures);
                }
                if let Some(rewritten) = go_rewrite_errors_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_sort_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_cmp_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_strings_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_strconv_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_time_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_url_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_atomic_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_slices_maps_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_slog_call(name, &next_args) {
                    return rewritten;
                }
            }

            if call_name.as_deref() == Some("recover") && next_args.is_empty() {
                return go_recover_iife_expr(env);
            }

            if call_name.as_deref() == Some("make") {
                if let Some(type_name) = next_args
                    .first()
                    .and_then(|arg| go_type_name_from_expr(&arg.value))
                {
                    if go_is_channel_type(&type_name) {
                        let capacity = next_args.get(1).map(|arg| arg.value.clone());
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(channels::channel_new_expr(capacity)),
                            type_name,
                        });
                    }
                    if go_is_slice_type(&type_name) {
                        let len_expr = next_args
                            .get(1)
                            .map(|arg| arg.value.clone())
                            .unwrap_or_else(|| Expression::int(0));
                        let init_expr = go_array_element_type(&type_name)
                            .map(|elem| go_zero_value_expr(&elem))
                            .unwrap_or_else(Expression::null);
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(go_array_make_expr(len_expr, init_expr)),
                            type_name,
                        });
                    }
                    if go_is_map_type(&type_name) {
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(Expression::new(ExprKind::Object(Vec::new()))),
                            type_name,
                        });
                    }
                }
            }

            if call_name.as_deref() == Some("new") {
                if let Some(type_name) = next_args
                    .first()
                    .and_then(|arg| go_type_name_from_expr(&arg.value))
                {
                    return Expression::new(ExprKind::Unary {
                        op: UnaryOp::AddrOf,
                        expr: Box::new(go_zero_value_for_type(&type_name, env)),
                    });
                }
            }

            if call_name.as_deref() == Some("strconv.Atoi") && next_args.len() == 1 {
                return Expression::new(ExprKind::Tuple(vec![
                    go_builtin_call("__go_to_int", vec![next_args[0].value.clone()]),
                    Expression::null(),
                ]));
            }

            if next_args.len() == 1 {
                if let Some(type_name) = call_name
                    .as_deref()
                    .filter(|name| go_is_type_conversion_target(name, env, signatures))
                {
                    return go_normalize_type_conversion(
                        type_name,
                        next_args[0].value.clone(),
                        env,
                        signatures,
                    );
                }
            }

            if call_name.as_deref() == Some("copy") && next_args.len() >= 2 {
                let target = next_args[0].value.clone();
                let source = next_args[1].value.clone();
                return go_lower_copy_expr(
                    target,
                    source,
                    go_expr_type_hint(&next_args[0].value, env, signatures),
                    go_expr_type_hint(&next_args[1].value, env, signatures),
                    state,
                );
            }

            if call_name.as_deref() == Some("append") && !next_args.is_empty() {
                let mut result = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(next_args[0].value.clone()),
                    right: Box::new(Expression::new(ExprKind::Array(Vec::new()))),
                });
                for arg in next_args.iter().skip(1) {
                    let rhs = if arg.spread {
                        arg.value.clone()
                    } else {
                        Expression::new(ExprKind::Array(vec![ArrayElement {
                            key: None,
                            value: arg.value.clone(),
                            spread: false,
                            by_ref: false,
                        }]))
                    };
                    result = go_member_call(result, "concat", vec![rhs]);
                }
                return result;
            }

            if call_name.as_deref() == Some("len") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return channels::channel_len_expr(next_args[0].value.clone());
                }
            }

            if call_name.as_deref() == Some("cap") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return channels::channel_cap_expr(next_args[0].value.clone());
                }
                if let Some(cap_expr) = go_expr_capacity_hint(&next_args[0].value, env) {
                    return cap_expr;
                }
            }

            if call_name.as_deref() == Some("strings.Replace")
                && next_args.len() == 4
                && go_is_neg_one_expr(&next_args[3].value)
            {
                next_args.pop();
            }

            if call_name.as_deref() == Some("strings.Fields") && next_args.len() == 1 {
                let trimmed = go_member_call(next_args[0].value.clone(), "trim", Vec::new());
                return go_builtin_call(
                    "__go_regex_split_pat_first",
                    vec![Expression::string("\\s+"), trimmed],
                );
            }

            if call_name.as_deref() == Some("close") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return channels::channel_close_expr(next_args[0].value.clone());
                }
            }

            if call_name.as_deref() == Some("__go_type_assert") && next_args.len() == 2 {
                if let Some(type_name) = go_type_name_from_expr(&next_args[1].value) {
                    return go_type_assert_value_expr(next_args[0].value.clone(), &type_name);
                }
            }

            Expression::new(ExprKind::Call {
                callee: Box::new(next_callee),
                args: next_args,
                optional: *optional,
            })
        }
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .as_ref()
                        .map(|key| normalize_go_expr(key, env, signatures, state)),
                    value: normalize_go_expr(&element.value, env, signatures, state),
                    spread: element.spread,
                    by_ref: element.by_ref,
                })
                .collect(),
        )),
        ExprKind::Object(props) => Expression::new(ExprKind::Object(
            props
                .iter()
                .map(|prop| match prop {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: normalize_go_expr(key, env, signatures, state),
                        value: normalize_go_expr(value, env, signatures, state),
                    },
                    ObjectProperty::Spread(value) => {
                        ObjectProperty::Spread(normalize_go_expr(value, env, signatures, state))
                    }
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: normalize_go_expr(key, env, signatures, state),
                        value: normalize_go_expr(value, env, signatures, state),
                    },
                    _ => prop.clone(),
                })
                .collect(),
        )),
        ExprKind::Cast { expr, type_name } => {
            let normalized_expr = normalize_go_expr(expr, env, signatures, state);
            go_normalize_typed_composite_expr(normalized_expr, type_name, env)
        }
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .iter()
                .map(|value| normalize_go_expr(value, env, signatures, state))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => {
            let mut lambda_env = GoNormalizeEnv {
                value_types: env.value_types.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                slice_caps: env.slice_caps.clone(),
                slice_views: env.slice_views.clone(),
                struct_infos: env.struct_infos.clone(),
                named_types: env.named_types.clone(),
                type_names: env.type_names.clone(),
                return_type: None,
                panic_value_name: env.panic_value_name.clone(),
                has_panic_name: env.has_panic_name.clone(),
                in_defer_name: env.in_defer_name.clone(),
                recover_fn_name: env.recover_fn_name.clone(),
                owns_panic_state: false,
            };
            for param in params {
                if let Some(type_hint) = param.type_hint.as_ref() {
                    lambda_env
                        .value_types
                        .insert(param.name.clone(), type_hint.clone());
                }
                if let Some(type_hint) = param
                    .type_hint
                    .as_deref()
                    .filter(|hint| go_is_fixed_array_type(hint))
                {
                    lambda_env
                        .fixed_arrays
                        .insert(param.name.clone(), type_hint.to_string());
                }
            }
            let next_body = match body {
                LambdaBody::Expr(expr) => LambdaBody::Expr(Box::new(normalize_go_expr(
                    expr,
                    &lambda_env,
                    signatures,
                    state,
                ))),
                LambdaBody::Block(stmts) => LambdaBody::Block(normalize_go_function_body(
                    stmts,
                    &mut lambda_env,
                    signatures,
                    state,
                )),
            };
            Expression::new(ExprKind::Lambda {
                params: params.clone(),
                body: next_body,
                is_async: *is_async,
                captures: captures.clone(),
            })
        }
        _ => expr.clone(),
    }
}

fn lower_go_fixed_array_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_type = go_expr_type_hint(&iter, env, signatures).unwrap_or_default();
    let iter_name = fresh_go_temp(state, "__go_range_iter");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: (!iter_type.is_empty()).then(|| iter_type.clone()),
            init: Some(go_wrap_fixed_array_copy(iter, env, signatures)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    if !iter_type.is_empty() {
        body_env
            .fixed_arrays
            .insert(iter_name.clone(), iter_type.clone());
    }

    let mut lowered_body = Vec::new();
    match key {
        Some(key_name) => {
            if key_name != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(key_name.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: go_array_element_type(&iter_type),
                            init: Some(Expression::new(ExprKind::Index {
                                object: Box::new(Expression::ident(&iter_name)),
                                index: Box::new(Expression::ident(&index_name)),
                                null_safe: false,
                            })),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
        None => {
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
    }

    for stmt in body {
        lowered_body.extend(normalize_go_statement(
            stmt,
            &mut body_env,
            signatures,
            state,
        ));
    }

    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(go_builtin_call("len", vec![Expression::ident(&iter_name)])),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&index_name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&index_name)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: lowered_body,
    });

    vec![Statement::new(StmtKind::Block(vec![iter_decl, for_stmt]))]
}

fn lower_go_string_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_name = fresh_go_temp(state, "__go_range_str");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: Some("string".to_string()),
            init: Some(iter),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    body_env
        .value_types
        .insert(iter_name.clone(), "string".to_string());

    let mut lowered_body = Vec::new();
    match key {
        Some(key_name) => {
            if key_name != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(key_name.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(go_member_call(
                                Expression::ident(&iter_name),
                                "charCodeAt",
                                vec![Expression::ident(&index_name)],
                            )),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
        None => {
            if var != "_" {
                lowered_body.extend(normalize_go_statement(
                    &Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(var.to_string()),
                            type_hint: Some("int".to_string()),
                            init: Some(Expression::ident(&index_name)),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    &mut body_env,
                    signatures,
                    state,
                ));
            }
        }
    }

    for stmt in body {
        lowered_body.extend(normalize_go_statement(
            stmt,
            &mut body_env,
            signatures,
            state,
        ));
    }

    let for_stmt = Statement::new(StmtKind::For {
        init: Some(Box::new(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(go_builtin_call("len", vec![Expression::ident(&iter_name)])),
        })),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident(&index_name)),
            value: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::ident(&index_name)),
                right: Box::new(Expression::int(1)),
            })),
        })),
        body: lowered_body,
    });

    vec![Statement::new(StmtKind::Block(vec![iter_decl, for_stmt]))]
}

fn fresh_go_temp(state: &mut GoNormalizeState, prefix: &str) -> String {
    let name = format!("{}{}", prefix, state.next_temp);
    state.next_temp += 1;
    name
}

fn go_wrap_fixed_array_copy(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_fixed_array(&expr, env, signatures) && go_requires_fixed_array_copy(&expr) {
        go_builtin_call("__go_fixed_array_clone", vec![expr])
    } else {
        expr
    }
}

fn go_requires_fixed_array_copy(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
    )
}

fn go_builtin_call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args
            .into_iter()
            .map(|value| Argument {
                value,
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect(),
        optional: false,
    })
}

/// Build a slice/array literal AST node from a list of element expressions.
fn go_array_of(elems: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        elems
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn go_arg_value(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|a| a.value.clone())
        .unwrap_or_else(Expression::null)
}

/// Rewrite `errors.*` / `fmt.Errorf` package calls into calls to the injected
/// runtime prelude helpers. `errors.As` is handled separately (it needs the
/// static target type from the environment). Returns None for anything else.
fn go_rewrite_errors_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "errors.New" => Some(go_builtin_call(
            "__go_new_error",
            vec![go_arg_value(args, 0), Expression::null(), Expression::null()],
        )),
        "errors.Unwrap" => Some(go_builtin_call(
            "__go_errors_unwrap",
            vec![go_arg_value(args, 0)],
        )),
        "errors.Is" => Some(go_builtin_call(
            "__go_errors_is",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        "errors.Join" => {
            // `errors.Join(a, b, ...)` collects its variadic args into a slice.
            // `errors.Join(errs...)` already passes a slice — forward it.
            if args.len() == 1 && args[0].spread {
                Some(go_builtin_call("__go_errors_join", vec![args[0].value.clone()]))
            } else {
                let elems: Vec<Expression> = args.iter().map(|a| a.value.clone()).collect();
                Some(go_builtin_call("__go_errors_join", vec![go_array_of(elems)]))
            }
        }
        "fmt.Errorf" => go_rewrite_errorf(args),
        _ => None,
    }
}

/// Rewrite `fmt.Errorf(format, args...)` into a `__go_new_error(msg, wrap, errs)`
/// construction. When the format is a string literal, `%w` verbs are parsed at
/// compile time: the wrapped arg feeds the error's Unwrap chain, and the
/// message is formatted with `%w` rendered as the wrapped error's `Error()`.
fn go_rewrite_errorf(args: &[Argument]) -> Option<Expression> {
    let fmt_arg = args.first()?;
    let format_args: Vec<Expression> = args.iter().skip(1).map(|a| a.value.clone()).collect();

    let ExprKind::Lit(Literal::Str(fmt)) = &fmt_arg.value.kind else {
        // Non-literal format: format everything, no wrap tracking.
        let mut sprintf_args = vec![fmt_arg.value.clone()];
        sprintf_args.extend(format_args);
        let msg = go_builtin_call("__go_sprintf", sprintf_args);
        return Some(go_builtin_call(
            "__go_new_error",
            vec![msg, Expression::null(), Expression::null()],
        ));
    };

    let (newfmt, wrap_positions) = go_parse_errorf_format(fmt);

    // Build the sprintf argument list; render each `%w` arg via its Error().
    let mut sprintf_args = vec![Expression::string(&newfmt)];
    for (i, a) in format_args.iter().enumerate() {
        if wrap_positions.contains(&i) {
            if matches!(a.kind, ExprKind::Lit(Literal::Null)) {
                sprintf_args.push(Expression::string(""));
            } else {
                sprintf_args.push(go_member_call(a.clone(), "Error", vec![]));
            }
        } else {
            sprintf_args.push(a.clone());
        }
    }
    let msg = go_builtin_call("__go_sprintf", sprintf_args);

    let non_nil_wraps: Vec<Expression> = wrap_positions
        .iter()
        .filter_map(|&p| format_args.get(p).cloned())
        .filter(|e| !matches!(e.kind, ExprKind::Lit(Literal::Null)))
        .collect();

    let (wrap, errs) = match non_nil_wraps.len() {
        0 => (Expression::null(), Expression::null()),
        1 => (non_nil_wraps.into_iter().next().unwrap(), Expression::null()),
        _ => (Expression::null(), go_array_of(non_nil_wraps)),
    };

    Some(go_builtin_call("__go_new_error", vec![msg, wrap, errs]))
}

/// Rewrite `errors.As(err, &target)` into a call to the `__go_errors_as`
/// prelude helper with a type-match predicate and an assignment closure built
/// from the static target type. `errors.As` is reflection-shaped (generic over
/// the target type), so the type-specific part is synthesized here rather than
/// in the generic helper.
fn go_rewrite_errors_as(
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let err = go_arg_value(args, 0);
    let target_arg = go_arg_value(args, 1);

    // errors.As(err, nil) is always false.
    if matches!(target_arg.kind, ExprKind::Lit(Literal::Null)) {
        return Expression::bool(false);
    }

    // Extract the pointed-to target lvalue and its static type from `&target`.
    let (target_expr, target_type) = match &target_arg.kind {
        ExprKind::RefOf(place) => {
            let expr = go_place_expr(place);
            let ty = go_expr_type_hint(&expr, env, signatures);
            (expr, ty)
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => {
            let ty = go_expr_type_hint(expr, env, signatures);
            ((**expr).clone(), ty)
        }
        _ => return Expression::bool(false),
    };

    let Some(target_type) = target_type else {
        return Expression::bool(false);
    };
    let target_type = target_type
        .trim()
        .trim_start_matches('*')
        .trim()
        .to_string();

    let x = "__go_as_x";
    let match_closure = Expression::new(ExprKind::Lambda {
        params: vec![go_error_param(x)],
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(Expression::new(
            ExprKind::IsType {
                expr: Box::new(Expression::ident(x)),
                type_name: target_type.clone(),
            },
        ))))]),
        is_async: false,
        captures: Vec::new(),
    });
    let assign_closure = Expression::new(ExprKind::Lambda {
        params: vec![go_error_param(x)],
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Assign {
            targets: vec![target_expr],
            value: go_type_assert_value_expr(Expression::ident(x), &target_type),
        })]),
        is_async: false,
        captures: Vec::new(),
    });

    go_builtin_call("__go_errors_as", vec![err, match_closure, assign_closure])
}

/// Rewrite composite `strings.*` calls to the injected strings-prelude helpers.
fn go_rewrite_strings_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "strings.TrimPrefix" => "__go_strings_TrimPrefix",
        "strings.TrimSuffix" => "__go_strings_TrimSuffix",
        "strings.CutPrefix" => "__go_strings_CutPrefix",
        "strings.CutSuffix" => "__go_strings_CutSuffix",
        "strings.Cut" => "__go_strings_Cut",
        "strings.Replace" if args.len() == 4 => "__go_strings_Replace",
        "strings.ContainsRune" => "__go_strings_ContainsRune",
        "strings.ContainsAny" => "__go_strings_ContainsAny",
        "strings.ContainsFunc" => "__go_strings_ContainsFunc",
        "strings.IndexByte" => "__go_strings_IndexByte",
        "strings.IndexRune" => "__go_strings_IndexRune",
        "strings.IndexAny" => "__go_strings_IndexAny",
        "strings.IndexFunc" => "__go_strings_IndexFunc",
        "strings.LastIndexByte" => "__go_strings_LastIndexByte",
        "strings.LastIndexAny" => "__go_strings_LastIndexAny",
        "strings.LastIndexFunc" => "__go_strings_LastIndexFunc",
        "strings.TrimLeft" => "__go_strings_TrimLeft",
        "strings.TrimRight" => "__go_strings_TrimRight",
        "strings.Trim" => "__go_strings_TrimCutset",
        "strings.EqualFold" => "__go_strings_EqualFold",
        "strings.Map" => "__go_strings_Map",
        "strings.Fields" => "__go_strings_Fields",
        "strings.FieldsFunc" => "__go_strings_FieldsFunc",
        "strings.SplitN" => "__go_strings_SplitN",
        "strings.SplitAfter" => "__go_strings_SplitAfter",
        "strings.SplitAfterN" => "__go_strings_SplitAfterN",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

/// Rewrite `strconv.*` conversions. Parse functions return a `(value, error)`
/// tuple; the string-based helpers route to the strconv prelude.
fn go_rewrite_strconv_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let tuple_with_nil = |value: Expression| {
        Expression::new(ExprKind::Tuple(vec![value, Expression::null()]))
    };
    match call_name {
        "strconv.ParseBool" => Some(go_builtin_call("__go_strconv_ParseBool", vec![arg(0)])),
        "strconv.CanBackquote" => Some(go_builtin_call("__go_strconv_CanBackquote", vec![arg(0)])),
        // strconv.FormatBool(b) → b ? "true" : "false"
        "strconv.FormatBool" => Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(arg(0)),
            then: Box::new(Expression::string("true")),
            else_: Box::new(Expression::string("false")),
        })),
        // strconv.Atoi(s) → (int(s), nil)
        "strconv.Atoi" => Some(tuple_with_nil(go_builtin_call("__go_to_int", vec![arg(0)]))),
        // strconv.ParseInt/ParseUint(s, base, bits) → (parseInt(s, base), nil)
        "strconv.ParseInt" | "strconv.ParseUint" => {
            let base = if args.len() >= 2 { arg(1) } else { Expression::int(10) };
            Some(tuple_with_nil(go_builtin_call(
                "__go_parse_int",
                vec![arg(0), base],
            )))
        }
        // strconv.ParseFloat(s, bits) → (parseFloat(s), nil)
        "strconv.ParseFloat" => Some(tuple_with_nil(go_builtin_call(
            "__go_parse_float",
            vec![arg(0)],
        ))),
        _ => None,
    }
}

/// Rewrite `time.*` constructor calls to the injected time-prelude helpers.
fn go_rewrite_time_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "time.Unix" => "__go_time_Unix",
        "time.Date" => "__go_time_Date",
        "time.Now" => "__go_time_Now",
        "time.UnixMilli" => "__go_time_UnixMilli",
        "time.UnixMicro" => "__go_time_UnixMicro",
        "time.FixedZone" => "__go_time_FixedZone",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

/// Rewrite a `time.<Const>` member (non-call) to its runtime value. Durations
/// and layout strings come from `[namespace_constants]`; `time.UTC` builds the
/// UTC location.
fn go_rewrite_time_member(field: &str) -> Option<Expression> {
    match field {
        "UTC" | "Local" => Some(go_builtin_call(
            "__go_time_FixedZone",
            vec![Expression::string("UTC"), Expression::int(0)],
        )),
        _ => None,
    }
}

/// Rewrite `log/slog` package calls to the slog prelude helpers.
fn go_rewrite_slog_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let direct = |helper: &str| {
        Some(go_builtin_call(
            helper,
            args.iter().map(|a| a.value.clone()).collect(),
        ))
    };
    match call_name {
        "slog.NewTextHandler" => direct("__go_slog_NewTextHandler"),
        "slog.NewJSONHandler" => direct("__go_slog_NewJSONHandler"),
        "slog.New" => direct("__go_slog_New"),
        "slog.Default" => direct("__go_slog_Default"),
        "slog.Int" => direct("__go_slog_Int"),
        "slog.Int64" => direct("__go_slog_Int64"),
        "slog.String" => direct("__go_slog_String"),
        "slog.Bool" => direct("__go_slog_Bool"),
        "slog.Float64" => direct("__go_slog_Float64"),
        "slog.Duration" => direct("__go_slog_Duration"),
        "slog.Any" => direct("__go_slog_Any"),
        // slog.Group(key, attrs...) — variadic tail → slice.
        "slog.Group" => {
            let key = go_arg_value(args, 0);
            let attrs: Vec<Expression> = args.iter().skip(1).map(|a| a.value.clone()).collect();
            Some(go_builtin_call("__go_slog_Group", vec![key, go_array_of(attrs)]))
        }
        _ => None,
    }
}

/// Rewrite a `slog.<Const>` member to its prelude value (`slog.LevelInfo` etc.).
fn go_rewrite_slog_member(field: &str) -> Option<Expression> {
    let helper = match field {
        "LevelDebug" => "__go_slog_LevelDebug",
        "LevelInfo" => "__go_slog_LevelInfo",
        "LevelWarn" => "__go_slog_LevelWarn",
        "LevelError" => "__go_slog_LevelError",
        _ => return None,
    };
    Some(go_builtin_call(helper, vec![]))
}

/// Rewrite `slices.*` / `maps.*` calls to the slices/maps prelude helpers.
/// `Insert`/`Replace` collect their variadic tail into a slice.
fn go_rewrite_slices_maps_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let direct = |helper: &str, args: &[Argument]| {
        Some(go_builtin_call(
            helper,
            args.iter().map(|a| a.value.clone()).collect(),
        ))
    };
    match call_name {
        "slices.Contains" => direct("__go_slices_Contains", args),
        "slices.Index" => direct("__go_slices_Index", args),
        "slices.IndexFunc" => direct("__go_slices_IndexFunc", args),
        "slices.Equal" => direct("__go_slices_Equal", args),
        "slices.Compare" => direct("__go_slices_Compare", args),
        "slices.Clone" => direct("__go_slices_Clone", args),
        "slices.Compact" => direct("__go_slices_Compact", args),
        "slices.Delete" => direct("__go_slices_Delete", args),
        "slices.Grow" => direct("__go_slices_Grow", args),
        "slices.Clip" => direct("__go_slices_Clip", args),
        "slices.BinarySearch" => direct("__go_slices_BinarySearch", args),
        "slices.BinarySearchFunc" => direct("__go_slices_BinarySearchFunc", args),
        "maps.Clone" => direct("__go_maps_Clone", args),
        "maps.Copy" => direct("__go_maps_Copy", args),
        "maps.DeleteFunc" => direct("__go_maps_DeleteFunc", args),
        // slices.Insert(s, i, vals...) / Replace(s, i, j, vals...) — variadic tail → slice.
        "slices.Insert" => {
            let head: Vec<Expression> = args.iter().take(2).map(|a| a.value.clone()).collect();
            let tail: Vec<Expression> = args.iter().skip(2).map(|a| a.value.clone()).collect();
            let mut call_args = head;
            call_args.push(go_array_of(tail));
            Some(go_builtin_call("__go_slices_Insert", call_args))
        }
        "slices.Replace" => {
            let head: Vec<Expression> = args.iter().take(3).map(|a| a.value.clone()).collect();
            let tail: Vec<Expression> = args.iter().skip(3).map(|a| a.value.clone()).collect();
            let mut call_args = head;
            call_args.push(go_array_of(tail));
            Some(go_builtin_call("__go_slices_Replace", call_args))
        }
        _ => None,
    }
}

/// Rewrite `sync/atomic` function-style ops to the atomic prelude helpers. All
/// typed variants (`LoadInt64`, `LoadUint32`, …) map by operation prefix.
fn go_rewrite_atomic_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let rest = call_name.strip_prefix("atomic.")?;
    let helper = if rest.starts_with("Load") {
        "__go_atomic_Load"
    } else if rest.starts_with("Store") {
        "__go_atomic_Store"
    } else if rest.starts_with("Add") {
        "__go_atomic_Add"
    } else if rest.starts_with("Swap") {
        "__go_atomic_Swap"
    } else if rest.starts_with("CompareAndSwap") {
        "__go_atomic_CAS"
    } else {
        return None;
    };
    Some(go_builtin_call(
        helper,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

/// Rewrite `net/url` package functions to the injected url-prelude helpers.
fn go_rewrite_url_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "url.Parse" => Some(go_builtin_call("__go_url_Parse", vec![go_arg_value(args, 0)])),
        "url.ParseRequestURI" => Some(go_builtin_call(
            "__go_url_ParseRequestURI",
            vec![go_arg_value(args, 0)],
        )),
        "url.PathEscape" => Some(go_builtin_call(
            "__go_url_PathEscape",
            vec![go_arg_value(args, 0)],
        )),
        "url.PathUnescape" => Some(go_builtin_call(
            "__go_url_PathUnescape",
            vec![go_arg_value(args, 0)],
        )),
        "url.QueryEscape" => Some(go_builtin_call("__go_url_qesc", vec![go_arg_value(args, 0)])),
        "url.QueryUnescape" => Some(go_builtin_call("__go_url_unesc", vec![go_arg_value(args, 0)])),
        "url.User" => Some(go_builtin_call("__go_url_User", vec![go_arg_value(args, 0)])),
        "url.UserPassword" => Some(go_builtin_call(
            "__go_url_UserPassword",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        "url.JoinPath" => {
            let base = go_arg_value(args, 0);
            let elems: Vec<Expression> = args.iter().skip(1).map(|a| a.value.clone()).collect();
            Some(go_builtin_call(
                "__go_url_JoinPath",
                vec![base, go_array_of(elems)],
            ))
        }
        _ => None,
    }
}

/// Bind a Go stdlib type name to the runtime backing type its package prelude
/// defines — Go's equivalent of the `.NET` component types / libc surface. Used
/// so `url.Values{}` / `var q url.Values` resolve to the prelude's named type
/// (whose methods dispatch by type stamp).
fn go_stdlib_type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim() {
        "url.Values" => Some("__goValues"),
        "url.URL" => Some("__goURL"),
        "url.Userinfo" => Some("__goUser"),
        "bytes.Buffer" => Some("__goBuffer"),
        "slog.Level" => Some("__goLevel"),
        "slog.Attr" => Some("__goAttr"),
        "slog.Logger" => Some("__goSlogLogger"),
        "slog.Handler" => Some("__goSlogHandler"),
        "slog.HandlerOptions" => Some("__goHandlerOptions"),
        _ => None,
    }
}

/// Rewrite `cmp` package ordering helpers to plain comparisons.
fn go_rewrite_cmp_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let bin = |op: BinOp, l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    match call_name {
        // cmp.Less(a, b) → a < b
        "cmp.Less" => Some(bin(BinOp::Lt, go_arg_value(args, 0), go_arg_value(args, 1))),
        // cmp.Compare(a, b) → a < b ? -1 : (a > b ? 1 : 0)
        "cmp.Compare" => {
            let a = go_arg_value(args, 0);
            let b = go_arg_value(args, 1);
            let gt = Expression::new(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Gt, a.clone(), b.clone())),
                then: Box::new(Expression::int(1)),
                else_: Box::new(Expression::int(0)),
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(bin(BinOp::Lt, a, b)),
                then: Box::new(Expression::int(-1)),
                else_: Box::new(gt),
            }))
        }
        _ => None,
    }
}

/// Rewrite closure-based `sort.*` calls to the injected sort prelude helpers.
/// The index-relative comparator/swap closures are synthesized here because
/// they capture the target slice.
fn go_rewrite_sort_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let ii = "__go_sort_i";
    let jj = "__go_sort_j";
    match call_name {
        // sort.Search(n, f) — pass straight through.
        "sort.Search" => Some(go_builtin_call(
            "__go_sort_search",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        // sort.SearchInts/Strings/Float64s(a, x) — lower-bound: first i with a[i] >= x.
        "sort.SearchInts" | "sort.SearchStrings" | "sort.SearchFloat64s" => {
            let a = go_arg_value(args, 0);
            let x = go_arg_value(args, 1);
            let pred = go_lambda(
                vec![go_int_param(ii)],
                vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(go_index(a.clone(), Expression::ident(ii))),
                        right: Box::new(x),
                    },
                ))))],
            );
            Some(go_builtin_call(
                "__go_sort_search",
                vec![go_builtin_call("len", vec![a]), pred],
            ))
        }
        // sort.Slice/SliceStable(a, less) — insertion sort via index comparator
        // and a swap closure over the slice.
        "sort.Slice" | "sort.SliceStable" => {
            let a = go_arg_value(args, 0);
            let less = go_arg_value(args, 1);
            // Swap via a temp: `t := a[i]; a[i] = a[j]; a[j] = t`. A tuple
            // multi-assign to two index targets does not mutate a captured
            // slice, so keep it to single-target assignments.
            let tmp = "__go_sort_t";
            let swap = go_lambda(
                vec![go_int_param(ii), go_int_param(jj)],
                vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(tmp.to_string()),
                            type_hint: None,
                            init: Some(go_index(a.clone(), Expression::ident(ii))),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_index(a.clone(), Expression::ident(ii))],
                        value: go_index(a.clone(), Expression::ident(jj)),
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_index(a.clone(), Expression::ident(jj))],
                        value: Expression::ident(tmp),
                    }),
                ],
            );
            Some(go_builtin_call(
                "__go_sort_slice",
                vec![go_builtin_call("len", vec![a]), less, swap],
            ))
        }
        // sort.SliceIsSorted(a, less) — direct.
        "sort.SliceIsSorted" => {
            let a = go_arg_value(args, 0);
            let less = go_arg_value(args, 1);
            Some(go_builtin_call(
                "__go_sort_is_sorted",
                vec![go_builtin_call("len", vec![a]), less],
            ))
        }
        // sort.IntsAreSorted/Float64sAreSorted/StringsAreSorted(a) — ascending order.
        "sort.IntsAreSorted" | "sort.Float64sAreSorted" | "sort.StringsAreSorted" => {
            let a = go_arg_value(args, 0);
            let less = go_lambda(
                vec![go_int_param(ii), go_int_param(jj)],
                vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(go_index(a.clone(), Expression::ident(ii))),
                        right: Box::new(go_index(a.clone(), Expression::ident(jj))),
                    },
                ))))],
            );
            Some(go_builtin_call(
                "__go_sort_is_sorted",
                vec![go_builtin_call("len", vec![a]), less],
            ))
        }
        _ => None,
    }
}

/// A single `int`-typed lambda parameter named `name`.
fn go_int_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: Some("int".to_string()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

/// `obj[idx]` index expression.
fn go_index(obj: Expression, idx: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(obj),
        index: Box::new(idx),
        null_safe: false,
    })
}

/// A block-bodied closure with the given params.
fn go_lambda(params: Vec<Param>, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

/// A single `error`-typed lambda parameter named `name`.
fn go_error_param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: Some("error".to_string()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

/// Reconstruct an lvalue `Expression` from a `PlaceExpr`.
fn go_place_expr(place: &PlaceExpr) -> Expression {
    match place {
        PlaceExpr::Ident(name) => Expression::ident(name),
        PlaceExpr::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Deref(expr) => Expression::new(ExprKind::RefLoad(expr.clone())),
    }
}

/// Parse a Go `fmt` format string, returning the format with each `%w` verb
/// rewritten to `%s` (so the wrapped error renders via `Error()`), plus the
/// zero-based argument positions consumed by `%w` verbs.
fn go_parse_errorf_format(fmt: &str) -> (String, Vec<usize>) {
    let chars: Vec<char> = fmt.chars().collect();
    let mut out = String::new();
    let mut wraps = Vec::new();
    let mut arg_index = 0usize;
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c != '%' {
            out.push(c);
            i += 1;
            continue;
        }
        // `%%` is a literal percent — consumes no arg.
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            out.push_str("%%");
            i += 2;
            continue;
        }
        // Copy the verb spec (flags/width/precision) up to the verb letter.
        out.push('%');
        i += 1;
        while i < chars.len() {
            let vc = chars[i];
            if vc.is_ascii_alphabetic() {
                if vc == 'w' {
                    wraps.push(arg_index);
                    out.push('s');
                } else {
                    out.push(vc);
                }
                arg_index += 1;
                i += 1;
                break;
            }
            out.push(vc);
            i += 1;
        }
    }
    (out, wraps)
}

fn go_recover_iife_expr(env: &GoNormalizeEnv) -> Expression {
    let Some(panic_value_name) = env.panic_value_name.as_ref() else {
        return Expression::null();
    };
    let Some(has_panic_name) = env.has_panic_name.as_ref() else {
        return Expression::null();
    };
    let Some(in_defer_name) = env.in_defer_name.as_ref() else {
        return Expression::null();
    };

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(Expression::ident(in_defer_name)),
                    right: Box::new(Expression::ident(has_panic_name)),
                }),
                then_body: vec![
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(has_panic_name)],
                        value: Expression::bool(false),
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(panic_value_name)))),
                ],
                elifs: Vec::new(),
                else_body: Some(vec![Statement::new(StmtKind::Return(Some(
                    Expression::null(),
                )))]),
            })]),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn go_extract_panic_expr(expr: &Expression) -> Option<&Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Ident(name) = &callee.kind else {
        return None;
    };
    if name != "panic" || args.len() != 1 {
        return None;
    }
    Some(&args[0].value)
}

fn go_copy_count_expr(target: Expression, source: Expression) -> Expression {
    let target_len = go_builtin_call("len", vec![target]);
    let source_len = go_builtin_call("len", vec![source]);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(target_len.clone()),
            right: Box::new(source_len.clone()),
        })),
        then: Box::new(target_len),
        else_: Box::new(source_len),
    })
}

fn go_add_expr(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn go_materialize_slice_view(view: GoSliceViewInfo) -> Expression {
    let mut args = vec![view.start];
    if let Some(end) = view.end {
        args.push(end);
    }
    go_member_call(view.base, "slice", args)
}

fn go_slice_view_index_expr(view: GoSliceViewInfo, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(view.base),
        index: Box::new(go_add_expr(view.start, index)),
        null_safe: false,
    })
}

fn go_rewrite_slice_view_index(
    object: &Expression,
    index: Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    go_expr_slice_view(object, env).map(|view| go_slice_view_index_expr(view, index))
}

fn go_expr_slice_view(expr: &Expression, env: &GoNormalizeEnv) -> Option<GoSliceViewInfo> {
    match &expr.kind {
        ExprKind::Ident(name) => env.slice_views.get(name).cloned(),
        ExprKind::Call { callee, args, .. } => {
            let ExprKind::Member { object, field, .. } = &callee.kind else {
                return None;
            };
            if field != "slice" {
                return None;
            }

            let parent = go_expr_slice_view(object, env);
            let parent_start = parent
                .as_ref()
                .map(|view| view.start.clone())
                .unwrap_or_else(|| Expression::int(0));
            let start = go_add_expr(
                parent_start.clone(),
                args.first()
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::int(0)),
            );
            let end = if let Some(end_arg) = args.get(1) {
                Some(go_add_expr(parent_start, end_arg.value.clone()))
            } else {
                parent.as_ref().and_then(|view| view.end.clone())
            };

            Some(GoSliceViewInfo {
                base: parent
                    .map(|view| view.base)
                    .unwrap_or_else(|| object.as_ref().clone()),
                start,
                end,
            })
        }
        _ => None,
    }
}

fn go_slice_view_is_self_referential(view: &GoSliceViewInfo, name: &str) -> bool {
    matches!(&view.base.kind, ExprKind::Ident(base_name) if base_name == name)
}

fn go_lower_copy_expr(
    target: Expression,
    source: Expression,
    target_type: Option<String>,
    source_type: Option<String>,
    state: &mut GoNormalizeState,
) -> Expression {
    let target_name = fresh_go_temp(state, "__go_copy_dst");
    let source_name = fresh_go_temp(state, "__go_copy_src");
    let count_name = fresh_go_temp(state, "__go_copy_count");
    let index_name = fresh_go_temp(state, "__go_copy_idx");

    let count_expr = go_copy_count_expr(
        Expression::ident(&target_name),
        Expression::ident(&source_name),
    );
    let loop_cond = Expression::new(ExprKind::Binary {
        op: BinOp::Lt,
        left: Box::new(Expression::ident(&index_name)),
        right: Box::new(Expression::ident(&count_name)),
    });
    let loop_update = Expression::new(ExprKind::Assign {
        target: Box::new(Expression::ident(&index_name)),
        value: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(Expression::int(1)),
        })),
    });
    let loop_body = vec![Statement::new(StmtKind::Expr(Expression::new(
        ExprKind::Assign {
            target: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&target_name)),
                index: Box::new(Expression::ident(&index_name)),
                null_safe: false,
            })),
            value: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident(&source_name)),
                index: Box::new(Expression::ident(&index_name)),
                null_safe: false,
            })),
        },
    )))];

    let mut body = vec![
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(target_name.clone()),
                type_hint: target_type,
                init: Some(target),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(source_name.clone()),
                type_hint: source_type,
                init: Some(source),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(count_name.clone()),
                type_hint: Some("int".to_string()),
                init: Some(count_expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::For {
            init: None,
            cond: Some(loop_cond),
            update: Some(loop_update),
            body: loop_body,
        }),
    ];
    body.push(Statement::new(StmtKind::Return(Some(Expression::ident(
        &count_name,
    )))));

    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn go_expr_to_place(expr: &Expression) -> Option<PlaceExpr> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(PlaceExpr::Ident(name.clone())),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Some(PlaceExpr::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Some(PlaceExpr::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::RefLoad(expr) => Some(PlaceExpr::Deref(expr.clone())),
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        } => Some(PlaceExpr::Deref(expr.clone())),
        _ => None,
    }
}

fn normalize_go_lvalue_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) => Expression::ident(name),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => {
            let next_object = normalize_go_expr(object, env, signatures, state);
            let next_index = normalize_go_expr(index, env, signatures, state);
            go_rewrite_slice_view_index(&next_object, next_index.clone(), env).unwrap_or_else(
                || {
                    Expression::new(ExprKind::Index {
                        object: Box::new(next_object),
                        index: Box::new(next_index),
                        null_safe: *null_safe,
                    })
                },
            )
        }
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(normalize_go_lvalue_expr(target, env, signatures, state)),
            value: Box::new(normalize_go_expr(value, env, signatures, state)),
        }),
        _ => normalize_go_expr(expr, env, signatures, state),
    }
}

fn go_is_two_value_binding_pattern(pattern: &BindingPattern) -> bool {
    matches!(pattern, BindingPattern::Array(elems) if elems.len() == 2)
}

fn go_normalize_map_lookup_tuple_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Index { object, index, .. } = &expr.kind else {
        return None;
    };
    let value_type = go_map_index_value_type(expr, env, signatures)?;
    let next_object = normalize_go_expr(object, env, signatures, state);
    let next_index = normalize_go_expr(index, env, signatures, state);
    Some(Expression::new(ExprKind::Tuple(vec![
        go_build_map_read_expr(next_object.clone(), next_index.clone(), &value_type),
        go_map_has_expr(next_object, next_index),
    ])))
}

fn go_map_has_expr(object: Expression, index: Expression) -> Expression {
    go_builtin_call("__go_map_has", vec![object, index])
}

fn go_map_index_value_type(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    let ExprKind::Index { object, .. } = &expr.kind else {
        return None;
    };
    go_expr_type_hint(object, env, signatures).and_then(|type_name| go_map_value_type(&type_name))
}

fn go_build_map_read_expr(object: Expression, index: Expression, value_type: &str) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_map_has_expr(object.clone(), index.clone())),
        then: Box::new(Expression::new(ExprKind::Index {
            object: Box::new(object),
            index: Box::new(index),
            null_safe: false,
        })),
        else_: Box::new(go_zero_value_expr(value_type)),
    })
}

fn go_member_call(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(object),
            field: field.to_string(),
            null_safe: false,
        })),
        args: args
            .into_iter()
            .map(|value| Argument {
                value,
                name: None,
                by_ref: false,
                spread: false,
            })
            .collect(),
        optional: false,
    })
}

fn go_expr_is_fixed_array(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    go_expr_type_hint(expr, env, signatures)
        .as_deref()
        .is_some_and(go_is_fixed_array_type)
}

fn go_expr_type_hint(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => env
            .value_types
            .get(name)
            .cloned()
            .or_else(|| env.fixed_arrays.get(name).cloned()),
        ExprKind::Lit(Literal::Int(_)) => Some("int".to_string()),
        ExprKind::Lit(Literal::Float(_)) => Some("float64".to_string()),
        ExprKind::Lit(Literal::Bool(_)) => Some("bool".to_string()),
        ExprKind::Lit(Literal::Str(_)) => Some("string".to_string()),
        ExprKind::Cast { type_name, .. } => Some(type_name.clone()),
        ExprKind::RefOf(place) => {
            let pointee_type = match place.as_ref() {
                PlaceExpr::Ident(name) => env
                    .value_types
                    .get(name)
                    .cloned()
                    .or_else(|| env.fixed_arrays.get(name).cloned()),
                PlaceExpr::Member {
                    object,
                    field,
                    null_safe,
                } => go_expr_type_hint(
                    &Expression::new(ExprKind::Member {
                        object: object.clone(),
                        field: field.clone(),
                        null_safe: *null_safe,
                    }),
                    env,
                    signatures,
                ),
                PlaceExpr::Index {
                    object,
                    index,
                    null_safe,
                } => go_expr_type_hint(
                    &Expression::new(ExprKind::Index {
                        object: object.clone(),
                        index: index.clone(),
                        null_safe: *null_safe,
                    }),
                    env,
                    signatures,
                ),
                PlaceExpr::Deref(expr) => {
                    go_expr_type_hint(expr, env, signatures).map(|type_name| {
                        type_name
                            .trim()
                            .trim_start_matches('*')
                            .trim_start_matches('^')
                            .trim()
                            .to_string()
                    })
                }
            }?;
            Some(format!("*{}", pointee_type.trim()))
        }
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => go_expr_type_hint(expr, env, signatures)
            .map(|type_name| format!("*{}", type_name.trim())),
        ExprKind::Unary {
            op: UnaryOp::Deref,
            expr,
        }
        | ExprKind::RefLoad(expr) => go_expr_type_hint(expr, env, signatures).map(|type_name| {
            type_name
                .trim()
                .trim_start_matches('*')
                .trim_start_matches('^')
                .trim()
                .to_string()
        }),
        ExprKind::Member { object, field, .. } => go_expr_type_hint(object, env, signatures)
            .and_then(|type_name| {
                go_resolve_struct_member_type(&type_name, field, env, &mut HashSet::new())
            }),
        ExprKind::IsType { .. } => Some("bool".to_string()),
        ExprKind::Index { object, .. } => {
            go_expr_type_hint(object, env, signatures).and_then(|type_name| {
                if type_name == "string" {
                    Some("byte".to_string())
                } else {
                    go_array_element_type(&type_name).or_else(|| go_map_value_type(&type_name))
                }
            })
        }
        ExprKind::Assign { value, .. } => go_expr_type_hint(value, env, signatures),
        ExprKind::Ternary { then, else_, .. } => {
            let then_type = go_expr_type_hint(then, env, signatures);
            let else_type = go_expr_type_hint(else_, env, signatures);
            if then_type == else_type {
                then_type
            } else {
                then_type.or(else_type)
            }
        }
        ExprKind::Binary { op, left, right } => {
            let left_type = go_expr_type_hint(left, env, signatures);
            let right_type = go_expr_type_hint(right, env, signatures);
            match op {
                BinOp::Add
                | BinOp::Sub
                | BinOp::Mul
                | BinOp::IDiv
                | BinOp::Mod
                | BinOp::BitAnd
                | BinOp::BitOr
                | BinOp::BitXor
                | BinOp::Shl
                | BinOp::Shr => {
                    if left_type.as_deref().is_some_and(go_is_integer_type)
                        && right_type.as_deref().is_some_and(go_is_integer_type)
                    {
                        Some("int".to_string())
                    } else {
                        left_type.or(right_type)
                    }
                }
                BinOp::Div => {
                    if left_type.as_deref().is_some_and(go_is_integer_type)
                        && right_type.as_deref().is_some_and(go_is_integer_type)
                    {
                        Some("int".to_string())
                    } else {
                        Some("float64".to_string())
                    }
                }
                _ => None,
            }
        }
        ExprKind::Call { callee, args, .. } => match &callee.kind {
            ExprKind::Ident(name) if name == "__go_fixed_array_clone" => args
                .first()
                .and_then(|arg| go_expr_type_hint(&arg.value, env, signatures)),
            ExprKind::Ident(name) if name == "__go_fixed_array_equal" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_regex_split_pat_first" => {
                Some("[]string".to_string())
            }
            ExprKind::Ident(name) if name == "__go_map_has" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_to_int" => Some("int".to_string()),
            ExprKind::Ident(name) if name == "__go_str_from_char_code" => {
                Some("string".to_string())
            }
            ExprKind::Ident(name) if name == "__go_type_assert" => args
                .get(1)
                .and_then(|arg| go_type_name_from_expr(&arg.value)),
            ExprKind::Member { field, .. } if field == "charCodeAt" => Some("int".to_string()),
            ExprKind::Ident(name) => signatures.get(name).and_then(|sig| sig.return_type.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn go_expr_call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let object_name = go_expr_call_name(object)?;
            Some(format!("{}.{}", object_name, field))
        }
        _ => None,
    }
}

fn go_struct_lookup_name(type_name: &str) -> Option<String> {
    go_named_receiver_type(type_name)
}

fn go_resolve_struct_member_path(
    type_name: &str,
    member: &str,
    env: &GoNormalizeEnv,
    seen: &mut HashSet<String>,
) -> Option<Vec<String>> {
    let lookup = go_struct_lookup_name(type_name)?;
    if !seen.insert(lookup.clone()) {
        return None;
    }
    let info = env.struct_infos.get(&lookup)?;
    if info.member_names.contains(member) {
        return Some(vec![member.to_string()]);
    }
    for (embedded_name, embedded_type) in &info.embedded_fields {
        if let Some(mut tail) = go_resolve_struct_member_path(embedded_type, member, env, seen) {
            let mut path = vec![embedded_name.clone()];
            path.append(&mut tail);
            return Some(path);
        }
    }
    None
}

fn go_resolve_struct_member_type(
    type_name: &str,
    member: &str,
    env: &GoNormalizeEnv,
    seen: &mut HashSet<String>,
) -> Option<String> {
    let lookup = go_struct_lookup_name(type_name)?;
    if !seen.insert(lookup.clone()) {
        return None;
    }
    let info = env.struct_infos.get(&lookup)?;
    if let Some(type_name) = info.member_types.get(member) {
        return Some(type_name.clone());
    }
    for (_, embedded_type) in &info.embedded_fields {
        if let Some(type_name) = go_resolve_struct_member_type(embedded_type, member, env, seen) {
            return Some(type_name);
        }
    }
    None
}

fn go_rewrite_promoted_member_access(
    object: Expression,
    field: &str,
    null_safe: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = go_expr_type_hint(&object, env, signatures)?;
    let path = go_resolve_struct_member_path(&receiver_type, field, env, &mut HashSet::new())?;
    if path.len() <= 1 {
        return None;
    }

    let mut expr = object;
    for segment in path {
        expr = Expression::new(ExprKind::Member {
            object: Box::new(expr),
            field: segment,
            null_safe,
        });
    }
    Some(expr)
}

fn go_rewrite_named_type_method_call(
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    if !env.named_types.contains_key(&lookup) {
        return None;
    }
    let info = env.struct_infos.get(&lookup)?;
    if !info.method_names.contains(field) {
        return None;
    }

    let mut rewritten_args = Vec::with_capacity(args.len() + 1);
    rewritten_args.push(Argument::positional((**object).clone()));
    rewritten_args.extend(args.iter().cloned());

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&lookup)),
            field: field.clone(),
            null_safe: false,
        })),
        args: rewritten_args,
        optional,
    }))
}

fn go_is_function_type(type_name: &str) -> bool {
    type_name.trim().starts_with("func(")
}

fn go_rewrite_callable_field_member_call(
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    let info = env.struct_infos.get(&lookup)?;
    if info.method_names.contains(field) {
        return None;
    }
    let field_type = info.member_types.get(field)?;
    if !go_is_function_type(field_type) {
        return None;
    }

    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Sequence(vec![callee.clone()]))),
        args: args.to_vec(),
        optional,
    }))
}

fn go_normalize_typed_composite_expr(
    expr: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
) -> Expression {
    if let ExprKind::Array(elements) = &expr.kind {
        if let Some(lookup) = go_struct_lookup_name(type_name) {
            if let Some(info) = env.struct_infos.get(&lookup) {
                let mut props = Vec::new();
                for (index, field_name) in info.field_order.iter().enumerate() {
                    let value = elements
                        .get(index)
                        .map(|element| element.value.clone())
                        .or_else(|| {
                            info.member_types
                                .get(field_name)
                                .map(|field_type| go_zero_value_for_type(field_type, env))
                        });
                    if let Some(value) = value {
                        props.push(ObjectProperty::KeyValue {
                            key: Expression::string(field_name),
                            value,
                        });
                    }
                }
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Object(props))),
                    type_name: type_name.to_string(),
                });
            }
        }
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn go_is_neg_one_expr(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(-1)) => true,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        } => matches!(expr.kind, ExprKind::Lit(Literal::Int(1))),
        _ => false,
    }
}

fn go_decl_fixed_array_binding(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };
    let type_name = decl.type_hint.clone().or_else(|| {
        decl.init
            .as_ref()
            .and_then(|expr| go_expr_type_hint(expr, env, signatures))
    })?;
    go_is_fixed_array_type(&type_name).then(|| (name.clone(), type_name))
}

fn go_decl_binding_type(
    decl: &VarDeclarator,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<(String, String)> {
    let BindingPattern::Ident(name) = &decl.pattern else {
        return None;
    };

    decl.type_hint
        .clone()
        .or_else(|| {
            decl.init
                .as_ref()
                .and_then(|expr| go_expr_type_hint(expr, env, signatures))
        })
        .map(|type_name| (name.clone(), type_name))
}

fn go_single_named_binding_pattern(pattern: &BindingPattern) -> Option<BindingPattern> {
    let BindingPattern::Array(elements) = pattern else {
        return None;
    };

    if elements.len() != 1 {
        return None;
    }

    let mut bound_name = None;
    for element in elements {
        match element {
            ArrayPatternElem::Hole => return None,
            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) => {
                if bound_name.is_some() {
                    return None;
                }
                bound_name = Some(name.clone());
            }
            _ => return None,
        }
    }

    bound_name.map(BindingPattern::Ident)
}

fn go_is_fixed_array_type(type_name: &str) -> bool {
    go_array_head(type_name)
        .map(|(head, _)| !head.trim().is_empty())
        .unwrap_or(false)
}

fn go_is_integer_type(type_name: &str) -> bool {
    matches!(
        type_name.trim(),
        "int"
            | "int8"
            | "int16"
            | "int32"
            | "int64"
            | "uint"
            | "uint8"
            | "uint16"
            | "uint32"
            | "uint64"
            | "uintptr"
            | "byte"
            | "rune"
    )
}

fn go_is_float_type(type_name: &str) -> bool {
    matches!(type_name.trim(), "float32" | "float64")
}

fn go_is_builtin_conversion_type(type_name: &str) -> bool {
    go_is_integer_type(type_name)
        || go_is_float_type(type_name)
        || matches!(type_name.trim(), "string" | "bool")
}

fn go_is_type_conversion_target(
    type_name: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    (go_is_builtin_conversion_type(type_name) || env.type_names.contains(type_name))
        && !signatures.contains_key(type_name)
}

fn go_normalize_type_conversion(
    type_name: &str,
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if let Some(underlying) = env
        .named_types
        .get(type_name)
        .filter(|underlying| underlying.as_str() != type_name)
    {
        let normalized = go_normalize_type_conversion(underlying, expr, env, signatures);
        return Expression::new(ExprKind::Cast {
            expr: Box::new(normalized),
            type_name: type_name.to_string(),
        });
    }

    if go_is_integer_type(type_name) {
        let int_expr = go_builtin_call("__go_to_int", vec![expr]);
        if type_name == "int" {
            return int_expr;
        }
        return Expression::new(ExprKind::Cast {
            expr: Box::new(int_expr),
            type_name: type_name.to_string(),
        });
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(go_is_integer_type)
    {
        return go_builtin_call("__go_str_from_char_code", vec![expr]);
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn walk_package_clause(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_import(pair: Pair<Rule>) -> Result<Import, String> {
    let mut path = String::new();
    let mut alias: Option<String> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::import_spec => {
                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => {
                            alias = Some(spec_inner.as_str().to_string());
                        }
                        Rule::string_literal => {
                            path = unquote(spec_inner.as_str());
                        }
                        _ => {}
                    }
                }
            }
            Rule::string_literal => {
                path = unquote(inner.as_str());
            }
            _ => {}
        }
    }

    Ok(Import {
        kind: ImportKind::Simple { path, alias },
        span: Span::default(),
    })
}

fn unquote(s: &str) -> String {
    if s.len() < 2 {
        return s.to_string();
    }

    if s.starts_with('`') && s.ends_with('`') {
        return s[1..s.len() - 1].to_string();
    }

    if !((s.starts_with('"') && s.ends_with('"')) || (s.starts_with('\'') && s.ends_with('\''))) {
        return s.to_string();
    }

    let mut out = String::new();
    let mut chars = s[1..s.len() - 1].chars();
    while let Some(ch) = chars.next() {
        if ch != '\\' {
            out.push(ch);
            continue;
        }

        match chars.next() {
            Some('n') => out.push('\n'),
            Some('r') => out.push('\r'),
            Some('t') => out.push('\t'),
            Some('\\') => out.push('\\'),
            Some('"') => out.push('"'),
            Some('\'') => out.push('\''),
            Some(other) => out.push(other),
            None => out.push('\\'),
        }
    }

    out
}

fn walk_top_level(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    match pair.as_rule() {
        Rule::function_declaration => Ok(Some(walk_function_decl(pair)?)),
        Rule::method_declaration => Ok(Some(walk_method_decl(pair)?)),
        Rule::var_declaration => Ok(Some(walk_var_decl(pair)?)),
        Rule::const_declaration => Ok(Some(walk_const_decl(pair)?)),
        Rule::type_declaration => walk_type_decl(pair),
        Rule::declaration => {
            for inner in pair.into_inner() {
                return walk_top_level(inner);
            }
            Ok(None)
        }
        _ => Ok(None),
    }
}

// ── Function declarations ─────────────────────────────────────────────────────────────

fn walk_function_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
                named_results = sig.named_results;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    for param in named_results.iter().rev() {
        body_stmts.insert(
            0,
            go_named_result_marker_stmt(
                &param.name,
                param.type_hint.as_deref().unwrap_or("object"),
            ),
        );
    }

    Ok(Statement::new(StmtKind::FunctionDecl {
        name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    }))
}

fn walk_method_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut receiver_name = String::new();
    let mut receiver_type = String::new();
    let mut receiver_owner = String::new();
    let mut method_name = String::new();
    let mut params = Vec::new();
    let mut body_stmts = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::receiver => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::ident_name => receiver_name = r_inner.as_str().to_string(),
                        Rule::type_annotation => {
                            receiver_type = walk_type(r_inner.clone());
                            receiver_owner = go_named_receiver_type(&receiver_type)
                                .unwrap_or_else(|| receiver_type.clone());
                        }
                        _ => {}
                    }
                }
            }
            Rule::ident_name => method_name = inner.as_str().to_string(),
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
                return_type = sig.return_type;
                named_results = sig.named_results;
            }
            Rule::function_body | Rule::block_statement => {
                body_stmts = walk_block(inner)?;
            }
            _ => {}
        }
    }

    for param in named_results.iter().rev() {
        body_stmts.insert(
            0,
            go_named_result_marker_stmt(
                &param.name,
                param.type_hint.as_deref().unwrap_or("object"),
            ),
        );
    }

    // Prepend receiver as first parameter
    params.insert(
        0,
        Param {
            name: if receiver_name.is_empty() {
                "self".to_string()
            } else {
                receiver_name
            },
            type_hint: Some(receiver_type.clone()),
            default: None,
            pass_by: PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        },
    );

    let method_stmt = Statement::new(StmtKind::FunctionDecl {
        name: method_name,
        params,
        return_type,
        body: body_stmts,
        modifiers: Modifiers::default(),
        handles: Vec::new(),
        is_async: false,
        is_generator: false,
        is_sub: false,
    });

    Ok(Statement::new(StmtKind::StructDecl {
        name: receiver_owner,
        interfaces: Vec::new(),
        members: vec![ClassMember::Method(Box::new(method_stmt))],
        visibility: Visibility::Public,
        decorators: Vec::new(),
    }))
}

fn walk_signature(pair: Pair<Rule>) -> Result<GoSignatureInfo, String> {
    let mut params = Vec::new();
    let mut return_type: Option<String> = None;
    let mut named_results = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::parameter_list => {
                params = walk_parameter_list(inner)?;
            }
            Rule::result => {
                for r_inner in inner.into_inner() {
                    match r_inner.as_rule() {
                        Rule::type_annotation => return_type = Some(walk_type(r_inner)),
                        Rule::parameter_list => {
                            let p = walk_parameter_list(r_inner)?;
                            named_results = p
                                .iter()
                                .filter(|param| !param.name.starts_with("__go_param_"))
                                .cloned()
                                .collect();
                            return_type = if p.len() == 1 {
                                p[0].type_hint.clone()
                            } else {
                                Some(format!("[{}]", p.len()))
                            };
                        }
                        _ => {}
                    }
                }
            }
            Rule::type_annotation => {
                return_type = Some(walk_type(inner));
            }
            _ => {}
        }
    }

    Ok(GoSignatureInfo {
        params,
        return_type,
        named_results,
    })
}

fn go_named_result_marker_stmt(name: &str, type_name: &str) -> Statement {
    Statement::new(StmtKind::Expr(go_builtin_call(
        "__go_named_result",
        vec![
            Expression::string(name),
            go_type_arg_expr(type_name.to_string()),
        ],
    )))
}

fn go_named_type_marker_stmt(name: &str, type_name: &str) -> Statement {
    Statement::new(StmtKind::Expr(go_builtin_call(
        "__go_named_type",
        vec![
            Expression::string(name),
            go_type_arg_expr(type_name.to_string()),
        ],
    )))
}

fn walk_parameter_list(pair: Pair<Rule>) -> Result<Vec<Param>, String> {
    let mut params = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::parameter_decl {
            let mut names = Vec::new();
            let mut type_hint: Option<String> = None;
            let mut is_rest = false;

            for p_inner in inner.into_inner() {
                match p_inner.as_rule() {
                    Rule::ident_name => names.push(p_inner.as_str().to_string()),
                    Rule::ident_list => {
                        for id in p_inner.into_inner() {
                            if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::type_annotation => type_hint = Some(walk_type(p_inner)),
                    Rule::variadic_parameter_type => {
                        is_rest = true;
                        for v_inner in p_inner.into_inner() {
                            if v_inner.as_rule() == Rule::type_annotation {
                                type_hint = Some(format!("[]{}", walk_type(v_inner)));
                            }
                        }
                    }
                    _ => {}
                }
            }

            if names.is_empty() && type_hint.is_some() {
                names.push(format!("__go_param_{}", params.len()));
            }

            for name in names {
                params.push(Param {
                    name,
                    type_hint: type_hint.clone(),
                    default: None,
                    pass_by: PassBy::Value,
                    is_rest,
                    is_kwargs: false,
                    is_optional: false,
                    is_nullable: false,
                });
            }
        }
    }
    Ok(params)
}

fn walk_type(pair: Pair<Rule>) -> String {
    if let Some(backing) = go_stdlib_type_binding(pair.as_str()) {
        return backing.to_string();
    }
    // Erase generic type arguments: a `Name[args]` instantiation becomes the
    // bare `Name` (Vybe is dynamically typed, so generics are type-erased).
    let mut inners = pair.clone().into_inner();
    if let Some(first) = inners.next() {
        if first.as_rule() == Rule::ident_name {
            if let Some(second) = inners.next() {
                if second.as_rule() == Rule::type_arguments {
                    return first.as_str().to_string();
                }
            }
        }
    }
    pair.as_str().to_string()
}

fn walk_block(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::block_statement | Rule::function_body => {
                stmts.append(&mut walk_block(inner)?);
            }
            Rule::statement_list => {
                for s in inner.into_inner() {
                    if s.as_rule() == Rule::statement {
                        stmts.push(walk_statement(s)?);
                    }
                }
            }
            Rule::statement => {
                stmts.push(walk_statement(inner)?);
            }
            _ => {}
        }
    }
    Ok(stmts)
}

// ── Variable declarations ─────────────────────────────────────────────────────────────

fn walk_var_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::var_spec | Rule::const_spec => {
                let (mut decls, _) = walk_var_spec(inner, VarDeclKind::Let)?;
                declarations.append(&mut decls);
            }
            Rule::var_group | Rule::const_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::var_spec || spec.as_rule() == Rule::const_spec {
                        let (mut decls, _) = walk_var_spec(spec, VarDeclKind::Let)?;
                        declarations.append(&mut decls);
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    }))
}

fn walk_const_decl(pair: Pair<Rule>) -> Result<Statement, String> {
    let mut declarations = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::const_spec => {
                let (mut decls, _, _) = walk_const_spec(inner, 0, None, None)?;
                declarations.append(&mut decls);
            }
            Rule::const_group => {
                let mut prev_inits: Option<Vec<Expression>> = None;
                let mut prev_type_hint: Option<String> = None;
                let mut iota_index = 0i64;
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::const_spec {
                        let (mut decls, next_inits, next_type_hint) = walk_const_spec(
                            spec,
                            iota_index,
                            prev_inits.clone(),
                            prev_type_hint.clone(),
                        )?;
                        declarations.append(&mut decls);
                        prev_inits = Some(next_inits);
                        prev_type_hint = next_type_hint;
                        iota_index += 1;
                    }
                }
            }
            _ => {}
        }
    }

    Ok(Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Const,
    }))
}

fn walk_const_spec(
    pair: Pair<Rule>,
    iota_index: i64,
    prev_inits: Option<Vec<Expression>>,
    prev_type_hint: Option<String>,
) -> Result<(Vec<VarDeclarator>, Vec<Expression>, Option<String>), String> {
    let (names, type_hint, init_values) = parse_go_var_spec(pair)?;
    let effective_type_hint = type_hint.or(prev_type_hint);
    let raw_inits = if init_values.is_empty() {
        prev_inits.unwrap_or_default()
    } else {
        init_values
    };
    let next_inits: Vec<Expression> = raw_inits
        .iter()
        .map(|expr| go_rewrite_iota_expr(expr, iota_index))
        .collect();

    if names.len() > 1 && !next_inits.is_empty() {
        let pattern = BindingPattern::Array(
            names
                .into_iter()
                .map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                })
                .collect(),
        );
        let init = if next_inits.len() == 1 {
            next_inits[0].clone()
        } else {
            Expression::new(ExprKind::Tuple(next_inits.clone()))
        };
        return Ok((
            vec![VarDeclarator {
                pattern,
                init: Some(init),
                type_hint: effective_type_hint.clone(),
                array_bounds: None,
                with_events: false,
            }],
            raw_inits,
            effective_type_hint,
        ));
    }

    let mut declarations = Vec::new();
    for name in names {
        if name == "_" {
            continue;
        }
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: next_inits.first().cloned(),
            type_hint: effective_type_hint.clone(),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, raw_inits, effective_type_hint))
}

fn parse_go_var_spec(
    pair: Pair<Rule>,
) -> Result<(Vec<String>, Option<String>, Vec<Expression>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident)
                        || id.as_str() == "_"
                    {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => init_values = walk_expression_list(inner)?,
            Rule::expression => init_values.push(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok((names, type_hint, init_values))
}

fn go_rewrite_iota_expr(expr: &Expression, iota_index: i64) -> Expression {
    match &expr.kind {
        ExprKind::Ident(name) if name == "iota" => Expression::int(iota_index),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op: *op,
            expr: Box::new(go_rewrite_iota_expr(expr, iota_index)),
        }),
        ExprKind::Binary { op, left, right } => Expression::new(ExprKind::Binary {
            op: *op,
            left: Box::new(go_rewrite_iota_expr(left, iota_index)),
            right: Box::new(go_rewrite_iota_expr(right, iota_index)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_iota_expr(cond, iota_index)),
            then: Box::new(go_rewrite_iota_expr(then, iota_index)),
            else_: Box::new(go_rewrite_iota_expr(else_, iota_index)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_iota_expr(callee, iota_index)),
            args: args
                .iter()
                .map(|arg| Argument {
                    value: go_rewrite_iota_expr(&arg.value, iota_index),
                    name: arg.name.clone(),
                    by_ref: arg.by_ref,
                    spread: arg.spread,
                })
                .collect(),
            optional: *optional,
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_iota_expr(object, iota_index)),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_iota_expr(object, iota_index)),
            index: Box::new(go_rewrite_iota_expr(index, iota_index)),
            null_safe: *null_safe,
        }),
        ExprKind::Cast { expr, type_name } => Expression::new(ExprKind::Cast {
            expr: Box::new(go_rewrite_iota_expr(expr, iota_index)),
            type_name: type_name.clone(),
        }),
        _ => expr.clone(),
    }
}

fn walk_var_spec(
    pair: Pair<Rule>,
    _kind: VarDeclKind,
) -> Result<(Vec<VarDeclarator>, Option<String>), String> {
    let mut names = Vec::new();
    let mut type_hint: Option<String> = None;
    let mut init_values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident)
                        || id.as_str() == "_"
                    {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::ident_name => names.push(inner.as_str().to_string()),
            Rule::type_annotation => type_hint = Some(walk_type(inner)),
            Rule::expression_list => {
                init_values = walk_expression_list(inner)?;
            }
            Rule::expression => {
                init_values.push(walk_expression(inner)?);
            }
            _ => {}
        }
    }

    if names.len() > 1 && !init_values.is_empty() {
        let pattern = BindingPattern::Array(
            names
                .into_iter()
                .map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                })
                .collect(),
        );
        let init = if init_values.len() == 1 {
            init_values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(init_values))
        };
        return Ok((
            vec![VarDeclarator {
                pattern,
                init: Some(init),
                type_hint,
                array_bounds: None,
                with_events: false,
            }],
            None,
        ));
    }

    let mut declarations = Vec::new();
    for name in names {
        if name == "_" {
            continue;
        }
        declarations.push(VarDeclarator {
            pattern: BindingPattern::Ident(name),
            init: init_values.first().cloned(),
            type_hint: type_hint.clone(),
            array_bounds: None,
            with_events: false,
        });
    }

    Ok((declarations, type_hint))
}

// ── Type declarations (struct, interface, type alias) ─────────────────────────────────

fn walk_type_decl(pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_spec => {
                let mut name = String::new();
                let mut type_str = String::new();

                for spec_inner in inner.into_inner() {
                    match spec_inner.as_rule() {
                        Rule::ident_name => name = spec_inner.as_str().to_string(),
                        Rule::type_annotation => {
                            if let Some(type_stmt) =
                                walk_named_type_annotation(name.clone(), spec_inner.clone())?
                            {
                                return Ok(Some(type_stmt));
                            }
                            type_str = walk_type(spec_inner);
                        }
                        Rule::struct_type => {
                            return Ok(Some(walk_struct_type(name, spec_inner)?));
                        }
                        Rule::interface_type => {
                            return Ok(Some(walk_interface_type(name, spec_inner)?));
                        }
                        _ => {}
                    }
                }

                // Keep named-type metadata in a marker statement so it does not
                // create a runtime binding that shadows generated type methods.
                if !type_str.is_empty() && !name.is_empty() {
                    return Ok(Some(go_named_type_marker_stmt(&name, &type_str)));
                }
            }
            Rule::type_group => {
                for spec in inner.into_inner() {
                    if spec.as_rule() == Rule::type_spec {
                        return walk_type_decl(spec.into());
                    }
                }
            }
            _ => {}
        }
    }
    Ok(None)
}

fn walk_struct_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::field_decl {
            let mut field_names = Vec::new();
            let mut field_type: Option<String> = None;

            for f_inner in inner.into_inner() {
                match f_inner.as_rule() {
                    Rule::ident_list => {
                        for id in f_inner.into_inner() {
                            if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                field_names.push(id.as_str().to_string());
                            }
                        }
                    }
                    Rule::ident_name => field_names.push(f_inner.as_str().to_string()),
                    Rule::type_annotation => field_type = Some(walk_type(f_inner)),
                    _ => {}
                }
            }

            if field_names.is_empty() {
                if let Some(type_name) = field_type.as_deref().and_then(go_embedded_field_name) {
                    field_names.push(type_name);
                }
            }

            for fname in field_names {
                members.push(ClassMember::Field {
                    name: fname,
                    type_hint: field_type.clone(),
                    init: None,
                    modifiers: Modifiers::default(),
                    with_events: false,
                    array_bounds: field_type.as_deref().and_then(go_fixed_array_bounds_exprs),
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::StructDecl {
        name,
        interfaces: Vec::new(),
        members,
        visibility: Visibility::Public,
        decorators: Vec::new(),
    }))
}

fn walk_interface_type(name: String, pair: Pair<Rule>) -> Result<Statement, String> {
    let mut members = Vec::new();

    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::interface_member {
            let mut method_name = String::new();
            let mut params = Vec::new();
            let mut return_type: Option<String> = None;

            for m_inner in inner.into_inner() {
                match m_inner.as_rule() {
                    Rule::ident_name => method_name = m_inner.as_str().to_string(),
                    Rule::signature => {
                        let sig = walk_signature(m_inner)?;
                        params = sig.params;
                        return_type = sig.return_type;
                    }
                    _ => {}
                }
            }

            if !method_name.is_empty() {
                members.push(InterfaceMember::Method {
                    name: method_name,
                    params,
                    return_type,
                    is_sub: false,
                    signature_source: None,
                });
            }
        }
    }

    Ok(Statement::new(StmtKind::InterfaceDecl {
        name,
        parents: Vec::new(),
        members,
        decorators: Vec::new(),
    }))
}

fn walk_named_type_annotation(name: String, pair: Pair<Rule>) -> Result<Option<Statement>, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::struct_type => return Ok(Some(walk_struct_type(name, inner)?)),
            Rule::interface_type => return Ok(Some(walk_interface_type(name, inner)?)),
            _ => {}
        }
    }
    Ok(None)
}

fn go_embedded_field_name(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    if trimmed.is_empty() {
        return None;
    }
    trimmed.rsplit('.').next().map(|name| name.to_string())
}

fn go_named_receiver_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    // Strip generic receiver type parameters: `Cell[T]` → `Cell`. A method
    // receiver is always a named type, so any `[...]` is a type-param list.
    let trimmed = trimmed.split('[').next().unwrap_or(trimmed).trim();
    if trimmed.is_empty() {
        return None;
    }
    Some(trimmed.to_string())
}

// ── Statements ─────────────────────────────────────────────────────────────────────────

fn walk_statement(pair: Pair<Rule>) -> Result<Statement, String> {
    let span = to_span(&pair);
    let rule = pair.as_rule();
    if rule == Rule::statement {
        if let Some(inner) = pair.into_inner().next() {
            let mut s = walk_statement(inner)?;
            if s.span.start_line == 0 {
                s.span = span;
            }
            return Ok(s);
        }
        return Ok(Statement::with_span(StmtKind::Empty, span));
    }

    let kind = match rule {
        Rule::empty_statement => StmtKind::Empty,
        Rule::block_statement => StmtKind::Block(walk_block(pair)?),
        Rule::expression_statement => {
            let expr = walk_expression(first_meaningful(pair)?)?;
            StmtKind::Expr(expr)
        }
        Rule::assignment_statement => walk_assignment(pair)?,
        Rule::short_var_declaration => walk_short_var_decl(pair)?,
        Rule::inc_dec_statement => walk_inc_dec(pair)?,
        Rule::var_declaration => walk_var_decl(pair)?.kind,
        Rule::const_declaration => walk_const_decl(pair)?.kind,
        Rule::if_statement => walk_if(pair)?,
        Rule::switch_statement => walk_switch(pair)?,
        Rule::select_statement => walk_select(pair)?,
        Rule::for_statement => walk_for(pair)?,
        Rule::return_statement => walk_return(pair)?,
        Rule::break_statement => {
            match pair.into_inner().find(|p| p.as_rule() == Rule::ident_name) {
                Some(lbl) => StmtKind::Break(BreakTarget::Label(lbl.as_str().to_string())),
                None => StmtKind::Break(BreakTarget::Implicit),
            }
        }
        Rule::continue_statement => {
            match pair.into_inner().find(|p| p.as_rule() == Rule::ident_name) {
                Some(lbl) => StmtKind::Continue(ContinueTarget::Label(lbl.as_str().to_string())),
                None => StmtKind::Continue(ContinueTarget::Implicit),
            }
        }
        Rule::fallthrough_statement => {
            StmtKind::Expr(Expression::ident(GO_FALLTHROUGH_MARKER))
        }
        Rule::goto_statement => StmtKind::GoTo(walk_goto(pair)?),
        Rule::labeled_statement => walk_labeled(pair)?,
        Rule::defer_statement => walk_defer_stmt(pair)?,
        Rule::go_statement => walk_go_stmt(pair)?,
        Rule::send_statement => walk_send_stmt(pair)?,
        _ => StmtKind::Empty,
    };
    Ok(Statement::with_span(kind, span))
}

fn walk_defer_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    Ok(StmtKind::Expr(go_builtin_call(
        "__go_defer",
        vec![walk_expression(first_meaningful(pair)?)?],
    )))
}

fn walk_go_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let expr = pair
        .into_inner()
        .find(|inner| inner.as_rule() == Rule::expression)
        .map(walk_expression)
        .transpose()?
        .unwrap_or_else(Expression::null);

    Ok(StmtKind::Expr(go_builtin_call(
        "__go_spawn",
        vec![go_wrap_spawn_expr(expr)],
    )))
}

fn walk_send_stmt(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression | Rule::primary => exprs.push(walk_expression(inner)?),
            _ => {}
        }
    }

    if exprs.len() == 2 {
        Ok(StmtKind::Expr(channels::channel_send_expr(
            exprs.remove(0),
            exprs.remove(0),
        )))
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_assignment(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut targets = Vec::new();
    let mut op = "=";
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression_list => {
                if targets.is_empty() {
                    targets = walk_expression_list(inner)?;
                } else {
                    values = walk_expression_list(inner)?;
                }
            }
            Rule::assign_op => op = inner.as_str(),
            _ => {}
        }
    }

    if op != "=" {
        // Compound assignment
        if targets.len() == 1 && values.len() == 1 {
            let target = targets[0].clone();
            let value = values[0].clone();
            let compound_op = match op {
                "+=" => Some(CompoundOp::Add),
                "-=" => Some(CompoundOp::Sub),
                "*=" => Some(CompoundOp::Mul),
                "/=" => Some(CompoundOp::Div),
                "%=" => Some(CompoundOp::Mod),
                "&=" => Some(CompoundOp::BitAnd),
                "|=" => Some(CompoundOp::BitOr),
                "^=" => Some(CompoundOp::BitXor),
                "<<=" => Some(CompoundOp::Shl),
                ">>=" => Some(CompoundOp::Shr),
                _ => None,
            };
            if let Some(compound_op) = compound_op {
                return Ok(StmtKind::CompoundAssign {
                    target,
                    op: compound_op,
                    value,
                });
            }
            if op == "&^=" {
                let rhs = Expression::new(ExprKind::Unary {
                    op: UnaryOp::BitNot,
                    expr: Box::new(value),
                });
                return Ok(StmtKind::Assign {
                    targets: vec![target.clone()],
                    value: Expression::new(ExprKind::Binary {
                        op: BinOp::BitAnd,
                        left: Box::new(target),
                        right: Box::new(rhs),
                    }),
                });
            }
        }
    }

    if targets.len() > 1 {
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(values))
        };
        if targets
            .iter()
            .all(|target| matches!(target.kind, ExprKind::Ident(_)))
            && !matches!(value.kind, ExprKind::Tuple(_))
        {
            let patterns = targets
                .iter()
                .map(|target| match &target.kind {
                    ExprKind::Ident(name) => {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name.clone()), None)
                    }
                    _ => ArrayPatternElem::Hole,
                })
                .collect();
            return Ok(StmtKind::Assign {
                targets: vec![Expression::new(ExprKind::Destructure(
                    DestructurePattern::Array(patterns),
                ))],
                value,
            });
        }
        return Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Tuple(targets))],
            value,
        });
    }

    if values.len() == 1 {
        Ok(StmtKind::Assign {
            targets,
            value: values.into_iter().next().unwrap(),
        })
    } else if !values.is_empty() {
        Ok(StmtKind::Assign {
            targets,
            value: Expression::new(ExprKind::Tuple(values)),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_short_var_decl(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut names = Vec::new();
    let mut values = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_list => {
                for id in inner.into_inner() {
                    if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                        names.push(id.as_str().to_string());
                    }
                }
            }
            Rule::expression_list => {
                values = walk_expression_list(inner)?;
            }
            _ => {}
        }
    }

    let mut declarations = Vec::new();
    if names.len() == 2 && values.len() == 1 {
        if let Some((expr, type_name)) = go_extract_type_assert_expr(&values[0]) {
            let pattern = BindingPattern::Array(
                names
                    .into_iter()
                    .map(|name| {
                        if name == "_" {
                            ArrayPatternElem::Hole
                        } else {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                        }
                    })
                    .collect(),
            );
            declarations.push(VarDeclarator {
                pattern,
                init: Some(Expression::new(ExprKind::Tuple(vec![
                    go_type_assert_value_expr(expr.clone(), &type_name),
                    Expression::new(ExprKind::IsType {
                        expr: Box::new(expr),
                        type_name,
                    }),
                ]))),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            });
            return Ok(StmtKind::VarDecl {
                declarations,
                kind: VarDeclKind::Let,
            });
        }
    }

    if names.len() > 1 && !values.is_empty() {
        let pattern = BindingPattern::Array(
            names
                .into_iter()
                .map(|name| {
                    if name == "_" {
                        ArrayPatternElem::Hole
                    } else {
                        ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                    }
                })
                .collect(),
        );
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else {
            Expression::new(ExprKind::Tuple(values))
        };
        declarations.push(VarDeclarator {
            pattern,
            init: Some(value),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        });
    } else {
        let value = if values.len() == 1 {
            values.into_iter().next().unwrap()
        } else if !values.is_empty() {
            Expression::new(ExprKind::Tuple(values))
        } else {
            Expression::new(ExprKind::Lit(Literal::Null))
        };

        for name in names {
            if name == "_" {
                continue;
            }
            declarations.push(VarDeclarator {
                pattern: BindingPattern::Ident(name),
                init: Some(value.clone()),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            });
        }
    }

    Ok(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_inc_dec(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut expr = None;
    let is_inc = !pair
        .as_str()
        .trim_end()
        .trim_end_matches(';')
        .trim_end()
        .ends_with("--");

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::primary => expr = Some(walk_primary(inner)?),
            _ => {}
        }
    }

    if let Some(target) = expr {
        Ok(StmtKind::CompoundAssign {
            target,
            op: if is_inc {
                CompoundOp::Add
            } else {
                CompoundOp::Sub
            },
            value: Expression::new(ExprKind::Lit(Literal::Int(1))),
        })
    } else {
        Ok(StmtKind::Empty)
    }
}

fn walk_if(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut cond = None;
    let mut then_body = Vec::new();
    let mut elifs = Vec::new();
    let mut else_body: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::expression | Rule::if_expression => {
                if cond.is_none() {
                    cond = Some(walk_expression(inner)?);
                }
            }
            Rule::block_statement => {
                if then_body.is_empty() {
                    then_body = walk_block(inner)?;
                }
            }
            Rule::else_clause => {
                for e_inner in inner.into_inner() {
                    match e_inner.as_rule() {
                        Rule::block_statement => else_body = Some(walk_block(e_inner)?),
                        Rule::if_statement => {
                            let elif = walk_if(e_inner)?;
                            match elif {
                                StmtKind::If {
                                    cond: c,
                                    then_body: t,
                                    elifs: nested_elifs,
                                    else_body: nested_else,
                                } => {
                                    elifs.push((c, t));
                                    elifs.extend(nested_elifs);
                                    else_body = nested_else;
                                }
                                StmtKind::Block(stmts) => else_body = Some(stmts),
                                _ => {}
                            }
                        }
                        _ => {}
                    }
                }
            }
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            _ => {}
        }
    }

    let if_stmt = Statement::new(StmtKind::If {
        cond: cond.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
        then_body,
        elifs,
        else_body,
    });

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![*pre, if_stmt]))
    } else {
        Ok(if_stmt.kind)
    }
}

/// Sentinel identifier a `fallthrough` statement is walked to, so `walk_switch`
/// can desugar it by inlining the following clause's body.
const GO_FALLTHROUGH_MARKER: &str = "__go_fallthrough__";

fn go_body_ends_with_fallthrough(body: &[Statement]) -> bool {
    matches!(
        body.last().map(|s| &s.kind),
        Some(StmtKind::Expr(Expression { kind: ExprKind::Ident(name), .. })) if name == GO_FALLTHROUGH_MARKER
    )
}

fn walk_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    match pair.as_rule() {
        Rule::switch_statement => {
            if let Some(inner) = pair.into_inner().next() {
                return walk_switch(inner);
            }
            return Ok(StmtKind::Empty);
        }
        Rule::type_switch_stmt => return walk_type_switch(pair),
        Rule::expr_switch_stmt => {}
        _ => return Ok(StmtKind::Empty),
    }

    let mut expr = None;
    let mut cases = Vec::new();
    let mut default: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;
    // Clauses in source order, so `fallthrough` can inline the next clause.
    let mut ordered: Vec<(Vec<CaseCondition>, Vec<Statement>)> = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::switch_short_var_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::switch_assignment_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::expression => expr = Some(walk_expression(inner)?),
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::expr_case_clause => {
                let mut conditions: Vec<CaseCondition> = Vec::new();
                let mut body = Vec::new();

                for c_inner in inner.into_inner() {
                    match c_inner.as_rule() {
                        Rule::expr_switch_case => {
                            for sc_inner in c_inner.into_inner() {
                                if sc_inner.as_rule() == Rule::expression_list {
                                    for expr in walk_expression_list(sc_inner)? {
                                        conditions.push(CaseCondition::Value(expr));
                                    }
                                } else if sc_inner.as_rule() == Rule::kw_default {
                                    // default case
                                }
                            }
                        }
                        Rule::statement_list => {
                            body = walk_statement_list(c_inner)?;
                        }
                        _ => {}
                    }
                }

                ordered.push((conditions, body));
            }
            _ => {}
        }
    }

    // Desugar `fallthrough`: a clause ending in the marker continues into the
    // next clause's (already-resolved) body.
    for i in (0..ordered.len()).rev() {
        if go_body_ends_with_fallthrough(&ordered[i].1) {
            let next_body = ordered.get(i + 1).map(|c| c.1.clone()).unwrap_or_default();
            let body = &mut ordered[i].1;
            body.pop(); // drop the marker
            body.extend(next_body);
        }
    }
    for (conditions, body) in ordered {
        if conditions.is_empty() {
            default = Some(body);
        } else {
            cases.push(SwitchCase { conditions, body });
        }
    }

    let switch_stmt = if expr.is_none() {
        let mut first_case: Option<(Expression, Vec<Statement>)> = None;
        let mut elifs = Vec::new();
        for case in cases {
            let cond = case
                .conditions
                .into_iter()
                .filter_map(|condition| match condition {
                    CaseCondition::Value(expr) => Some(expr),
                    _ => None,
                })
                .reduce(|left, right| {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Or,
                        left: Box::new(left),
                        right: Box::new(right),
                    })
                })
                .unwrap_or_else(|| Expression::bool(false));
            if first_case.is_none() {
                first_case = Some((cond, case.body));
            } else {
                elifs.push((cond, case.body));
            }
        }

        if let Some((cond, then_body)) = first_case {
            StmtKind::If {
                cond,
                then_body,
                elifs,
                else_body: default,
            }
        } else {
            StmtKind::Block(default.unwrap_or_default())
        }
    } else {
        StmtKind::Switch {
            expr: expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Bool(true)))),
            cases,
            default,
        }
    };

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![*pre, Statement::new(switch_stmt)]))
    } else {
        Ok(switch_stmt)
    }
}

fn walk_type_switch(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut binding_name: Option<String> = None;
    let mut switch_expr: Option<Expression> = None;
    let mut first_case: Option<(Expression, Vec<Statement>)> = None;
    let mut elifs = Vec::new();
    let mut default_body: Option<Vec<Statement>> = None;
    let mut pre_stmt: Option<Box<Statement>> = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::switch_short_var_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::switch_assignment_init => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::short_var_declaration => {
                pre_stmt = Some(Box::new(Statement::new(walk_short_var_decl(inner)?)));
            }
            Rule::expression_statement => {
                let expr = walk_expression(first_meaningful(inner)?)?;
                pre_stmt = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
            }
            Rule::assignment_statement => {
                pre_stmt = Some(Box::new(Statement::new(walk_assignment(inner)?)));
            }
            Rule::type_switch_guard => {
                for guard_inner in inner.into_inner() {
                    match guard_inner.as_rule() {
                        Rule::ident_name => binding_name = Some(guard_inner.as_str().to_string()),
                        Rule::primary | Rule::type_switch_subject => {
                            switch_expr = Some(walk_primary(guard_inner)?)
                        }
                        _ => {}
                    }
                }
            }
            Rule::type_case_clause => {
                let mut case_types = Vec::new();
                let mut body = Vec::new();
                for case_inner in inner.into_inner() {
                    match case_inner.as_rule() {
                        Rule::type_switch_case => {
                            for switch_case_inner in case_inner.into_inner() {
                                match switch_case_inner.as_rule() {
                                    Rule::type_list => {
                                        for ty in switch_case_inner.into_inner() {
                                            if ty.as_rule() == Rule::type_annotation {
                                                case_types.push(walk_type(ty));
                                            }
                                        }
                                    }
                                    Rule::kw_default => {}
                                    _ => {}
                                }
                            }
                        }
                        Rule::statement_list => body = walk_statement_list(case_inner)?,
                        _ => {}
                    }
                }

                if case_types.is_empty() {
                    default_body = Some(body);
                } else {
                    let expr = switch_expr.clone().unwrap_or_else(Expression::null);
                    let cond = go_type_switch_case_cond(expr.clone(), &case_types);
                    let case_body = go_type_switch_case_body(
                        body,
                        binding_name.as_deref(),
                        expr,
                        &case_types[0],
                    );
                    if first_case.is_none() {
                        first_case = Some((cond, case_body));
                    } else {
                        elifs.push((cond, case_body));
                    }
                }
            }
            _ => {}
        }
    }

    let type_switch_stmt = if let Some((cond, then_body)) = first_case {
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body: default_body,
        }
    } else {
        StmtKind::Block(default_body.unwrap_or_default())
    };

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![
            *pre,
            Statement::new(type_switch_stmt),
        ]))
    } else {
        Ok(type_switch_stmt)
    }
}

fn walk_select(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut arms: Vec<(Expression, Vec<Statement>)> = Vec::new();
    let mut default_body = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_clause => {
                for clause in inner.into_inner() {
                    match clause.as_rule() {
                        Rule::select_case_clause => {
                            if let Some(arm) = walk_select_case_clause(clause)? {
                                arms.push(arm);
                            }
                        }
                        Rule::select_default_clause => {
                            default_body = Some(walk_select_default_clause(clause)?)
                        }
                        _ => {}
                    }
                }
            }
            Rule::select_case_clause => {
                if let Some(arm) = walk_select_case_clause(inner)? {
                    arms.push(arm);
                }
            }
            Rule::select_default_clause => default_body = Some(walk_select_default_clause(inner)?),
            _ => {}
        }
    }

    let mut arm_iter = arms.into_iter();
    if let Some((cond, then_body)) = arm_iter.next() {
        Ok(StmtKind::If {
            cond,
            then_body,
            elifs: arm_iter.collect(),
            else_body: default_body,
        })
    } else {
        Ok(StmtKind::Block(default_body.unwrap_or_default()))
    }
}

fn go_select_receive_channel(expr: &Expression) -> Option<Expression> {
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__vybe_channel_receive") {
            return args.first().map(|arg| arg.value.clone());
        }
    }
    None
}

fn go_select_ready_cond(channel: Expression, is_send: bool) -> Expression {
    let left = channels::channel_len_expr(channel.clone());
    let right = if is_send {
        channels::channel_cap_expr(channel)
    } else {
        Expression::int(0)
    };

    Expression::new(ExprKind::Binary {
        op: if is_send { BinOp::Lt } else { BinOp::Gt },
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn walk_select_case_clause(
    pair: Pair<Rule>,
) -> Result<Option<(Expression, Vec<Statement>)>, String> {
    let mut prefix = Vec::new();
    let mut body = Vec::new();
    let mut cond = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_comm_clause => {
                let (comm_cond, mut comm_prefix) = walk_select_comm_clause(inner)?;
                cond = Some(comm_cond);
                prefix.append(&mut comm_prefix);
            }
            Rule::statement_list => body.extend(walk_statement_list(inner)?),
            _ => {}
        }
    }

    prefix.extend(body);
    Ok(cond.map(|cond| (cond, prefix)))
}

fn walk_select_default_clause(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement_list {
            return walk_statement_list(inner);
        }
    }
    Ok(Vec::new())
}

fn walk_select_comm_clause(pair: Pair<Rule>) -> Result<(Expression, Vec<Statement>), String> {
    let mut cond = None;
    let mut stmts = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_send_clause => {
                let mut exprs = Vec::new();
                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::expression | Rule::primary => exprs.push(walk_expression(part)?),
                        _ => {}
                    }
                }
                if exprs.len() == 2 {
                    cond = Some(go_select_ready_cond(exprs[0].clone(), true));
                    stmts.push(Statement::new(StmtKind::Expr(channels::channel_send_expr(
                        exprs.remove(0),
                        exprs.remove(0),
                    ))));
                }
            }
            Rule::select_receive_clause => {
                let mut names = Vec::new();
                let mut recv_expr = None;

                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::ident_list => {
                            for id in part.into_inner() {
                                if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                    names.push(id.as_str().to_string());
                                }
                            }
                        }
                        Rule::expression => recv_expr = Some(walk_expression(part)?),
                        _ => {}
                    }
                }

                if let Some(expr) = recv_expr {
                    if let Some(channel) = go_select_receive_channel(&expr) {
                        cond = Some(go_select_ready_cond(channel, false));
                    }
                    if names.is_empty() {
                        stmts.push(Statement::new(StmtKind::Expr(expr)));
                    } else {
                        stmts.push(go_short_var_decl_from_parts(names, expr));
                    }
                }
            }
            _ => {}
        }
    }

    Ok((cond.unwrap_or_else(|| Expression::bool(false)), stmts))
}

fn go_short_var_decl_from_parts(names: Vec<String>, value: Expression) -> Statement {
    let declarations = if names.len() > 1 {
        vec![VarDeclarator {
            pattern: BindingPattern::Array(
                names
                    .into_iter()
                    .map(|name| {
                        if name == "_" {
                            ArrayPatternElem::Hole
                        } else {
                            ArrayPatternElem::Pattern(BindingPattern::Ident(name), None)
                        }
                    })
                    .collect(),
            ),
            init: Some(Expression::new(ExprKind::Tuple(vec![
                value,
                Expression::bool(true),
            ]))),
            type_hint: None,
            array_bounds: None,
            with_events: false,
        }]
    } else {
        names
            .into_iter()
            .filter(|name| name != "_")
            .map(|name| VarDeclarator {
                pattern: BindingPattern::Ident(name),
                init: Some(value.clone()),
                type_hint: None,
                array_bounds: None,
                with_events: false,
            })
            .collect()
    };

    Statement::new(StmtKind::VarDecl {
        declarations,
        kind: VarDeclKind::Let,
    })
}

fn walk_statement_list(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    let mut stmts = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement {
            stmts.push(walk_statement(inner)?);
        }
    }
    Ok(stmts)
}

fn walk_for(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut init: Option<Box<Statement>> = None;
    let mut cond: Option<Expression> = None;
    let mut update: Option<Expression> = None;
    let mut body = Vec::new();
    let mut is_range = false;
    let mut range_vars = Vec::new();
    let mut range_iter = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::for_clause => {
                for fc_inner in inner.into_inner() {
                    match fc_inner.as_rule() {
                        Rule::for_short_var_nosemi => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::short_var_declaration => {
                            init = Some(Box::new(Statement::new(walk_short_var_decl(fc_inner)?)));
                        }
                        Rule::expression_statement => {
                            let expr = walk_expression(first_meaningful(fc_inner)?)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(StmtKind::Expr(expr))));
                            } else if update.is_none() {
                                update = Some(expr);
                            }
                        }
                        Rule::assignment_statement => {
                            let assign = walk_assignment(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(assign)));
                            } else if update.is_none() {
                                if let StmtKind::Assign { targets, value } = assign {
                                    if let Some(target) = targets.into_iter().next() {
                                        update = Some(Expression::new(ExprKind::Assign {
                                            target: Box::new(target),
                                            value: Box::new(value),
                                        }));
                                    }
                                }
                            }
                        }
                        Rule::inc_dec_statement => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if init.is_none() {
                                init = Some(Box::new(Statement::new(inc_dec)));
                            } else if update.is_none() {
                                if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                    let bin_op = match op {
                                        CompoundOp::Add => BinOp::Add,
                                        CompoundOp::Sub => BinOp::Sub,
                                        CompoundOp::Mul => BinOp::Mul,
                                        CompoundOp::Div => BinOp::Div,
                                        CompoundOp::Mod => BinOp::Mod,
                                        _ => BinOp::Add,
                                    };
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target.clone()),
                                        value: Box::new(Expression::new(ExprKind::Binary {
                                            op: bin_op,
                                            left: Box::new(target),
                                            right: Box::new(value),
                                        })),
                                    }));
                                }
                            }
                        }
                        Rule::for_inc_dec => {
                            let inc_dec = walk_inc_dec(fc_inner)?;
                            if let StmtKind::CompoundAssign { target, op, value } = inc_dec {
                                let bin_op = match op {
                                    CompoundOp::Add => BinOp::Add,
                                    CompoundOp::Sub => BinOp::Sub,
                                    CompoundOp::Mul => BinOp::Mul,
                                    CompoundOp::Div => BinOp::Div,
                                    CompoundOp::Mod => BinOp::Mod,
                                    _ => BinOp::Add,
                                };
                                update = Some(Expression::new(ExprKind::Assign {
                                    target: Box::new(target.clone()),
                                    value: Box::new(Expression::new(ExprKind::Binary {
                                        op: bin_op,
                                        left: Box::new(target),
                                        right: Box::new(value),
                                    })),
                                }));
                            }
                        }
                        Rule::for_assign_nosemi => {
                            let assign = walk_assignment(fc_inner)?;
                            if let StmtKind::Assign { targets, value } = assign {
                                if let Some(target) = targets.into_iter().next() {
                                    update = Some(Expression::new(ExprKind::Assign {
                                        target: Box::new(target),
                                        value: Box::new(value),
                                    }));
                                }
                            }
                        }
                        Rule::expression => {
                            if cond.is_none() {
                                cond = Some(walk_expression(fc_inner)?);
                            } else if update.is_none() {
                                update = Some(walk_expression(fc_inner)?);
                            }
                        }
                        Rule::block_statement => {
                            body = walk_block(fc_inner)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::range_clause => {
                is_range = true;
                for rc_inner in inner.into_inner() {
                    match rc_inner.as_rule() {
                        Rule::expression_list => {
                            for expr in walk_expression_list(rc_inner)? {
                                let name = if let ExprKind::Ident(id) = &expr.kind {
                                    id.clone()
                                } else {
                                    "_".to_string()
                                };
                                range_vars.push(BindingPattern::Ident(name));
                            }
                        }
                        Rule::ident_list => {
                            for id in rc_inner.into_inner() {
                                if matches!(id.as_rule(), Rule::ident_name | Rule::blank_ident) {
                                    range_vars.push(BindingPattern::Ident(id.as_str().to_string()));
                                }
                            }
                        }
                        Rule::expression | Rule::range_expression => {
                            range_iter = Some(walk_expression(rc_inner)?);
                        }
                        Rule::block_statement => {
                            body = walk_block(rc_inner)?;
                        }
                        _ => {}
                    }
                }
            }
            Rule::expression => {
                cond = Some(walk_expression(inner)?);
            }
            Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    if is_range {
        let var = if range_vars.len() > 1 {
            range_vars
                .get(1)
                .cloned()
                .unwrap_or_else(|| BindingPattern::Ident("_".to_string()))
        } else if range_vars.len() == 1 {
            BindingPattern::Ident("_".to_string())
        } else {
            range_vars
                .get(0)
                .cloned()
                .unwrap_or_else(|| BindingPattern::Ident("_".to_string()))
        };
        let var_name = match var {
            BindingPattern::Ident(name) => name,
            _ => "_".to_string(),
        };
        let key = if range_vars.len() > 1 {
            let key_pat = range_vars.get(0).cloned().unwrap();
            match key_pat {
                BindingPattern::Ident(name) => Some(name),
                _ => None,
            }
        } else if range_vars.len() == 1 {
            let key_pat = range_vars.get(0).cloned().unwrap();
            match key_pat {
                BindingPattern::Ident(name) => Some(name),
                _ => None,
            }
        } else {
            None
        };

        Ok(StmtKind::ForIn {
            var: var_name,
            key,
            iter: range_iter.unwrap_or_else(|| Expression::new(ExprKind::Array(Vec::new()))),
            body,
            of: true,
            else_body: None,
            is_async: false,
        })
    } else if init.is_none() && update.is_none() && cond.is_some() {
        Ok(StmtKind::While {
            cond: cond.unwrap(),
            body,
            else_body: None,
        })
    } else {
        Ok(StmtKind::For {
            init,
            cond,
            update,
            body,
        })
    }
}

fn walk_return(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut values = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression_list {
            values = walk_expression_list(inner)?;
        } else if inner.as_rule() == Rule::expression {
            values.push(walk_expression(inner)?);
        }
    }

    if values.len() == 1 {
        Ok(StmtKind::Return(Some(values.into_iter().next().unwrap())))
    } else if values.len() > 1 {
        let arr_elems: Vec<ArrayElement> = values
            .into_iter()
            .map(|v| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(StmtKind::Return(Some(Expression::new(ExprKind::Array(
            arr_elems,
        )))))
    } else {
        Ok(StmtKind::Return(None))
    }
}

fn walk_goto(pair: Pair<Rule>) -> Result<String, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::ident_name {
            return Ok(inner.as_str().to_string());
        }
    }
    Ok(String::new())
}

fn walk_labeled(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut label = String::new();
    let mut stmt = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => label = inner.as_str().to_string(),
            Rule::statement => stmt = Some(walk_statement(inner)?),
            _ => {}
        }
    }

    if let Some(s) = stmt {
        // Wrap the labeled statement so the compiler can route
        // `break <label>` / `continue <label>` to it.
        Ok(StmtKind::Labeled {
            label,
            body: Box::new(s),
        })
    } else {
        Ok(StmtKind::Label(label))
    }
}

// ── Expressions ─────────────────────────────────────────────────────────────────────────

fn walk_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    if matches!(
        pair.as_rule(),
        Rule::expression | Rule::if_expression | Rule::range_expression
    ) {
        let mut operands = Vec::new();
        let mut operators: Vec<String> = Vec::new();

        for inner in pair.into_inner() {
            match inner.as_rule() {
                Rule::unary_expression
                | Rule::if_unary_expression
                | Rule::range_unary_expression => operands.push(walk_unary_expression(inner)?),
                Rule::binary_op => {
                    let op = inner.as_str().to_string();
                    while operators
                        .last()
                        .is_some_and(|top| go_binary_precedence(top) >= go_binary_precedence(&op))
                    {
                        go_reduce_binary_expr(&mut operands, &mut operators)?;
                    }
                    operators.push(op);
                }
                _ => {}
            }
        }

        while !operators.is_empty() {
            go_reduce_binary_expr(&mut operands, &mut operators)?;
        }

        if let Some(result) = operands.pop() {
            return Ok(result);
        }
    } else if matches!(
        pair.as_rule(),
        Rule::unary_expression | Rule::if_unary_expression | Rule::range_unary_expression
    ) {
        return walk_unary_expression(pair);
    } else if matches!(
        pair.as_rule(),
        Rule::primary | Rule::if_primary | Rule::range_primary
    ) {
        return walk_primary(pair);
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_unary_expression(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut op = None;
    let mut operand = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::unary_op => op = Some(inner.as_str().to_string()),
            Rule::unary_expression | Rule::if_unary_expression | Rule::range_unary_expression => {
                operand = Some(walk_unary_expression(inner)?)
            }
            Rule::primary | Rule::if_primary | Rule::range_primary => {
                operand = Some(walk_primary(inner)?)
            }
            _ => {}
        }
    }

    if let Some(uop) = op {
        let un_op = match uop.as_str() {
            "-" => UnaryOp::Neg,
            "!" => UnaryOp::Not,
            "+" => UnaryOp::Pos,
            "^" => UnaryOp::BitNot,
            "*" => UnaryOp::Deref,
            "&" => UnaryOp::AddrOf,
            "<-" => {
                return Ok(channels::channel_receive_expr(
                    operand.unwrap_or_else(Expression::null),
                ));
            }
            _ => UnaryOp::Pos,
        };
        Ok(Expression::new(ExprKind::Unary {
            op: un_op,
            expr: Box::new(
                operand.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null))),
            ),
        }))
    } else {
        operand.ok_or_else(|| "Empty unary expression".to_string())
    }
}

fn walk_primary(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut base = None;
    let mut chain = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::operand | Rule::if_operand | Rule::range_operand => {
                base = Some(walk_operand(inner)?);
            }
            Rule::selector => {
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::ident_name {
                        chain.push(PrimaryChain::Member(s_inner.as_str().to_string()));
                    }
                }
            }
            Rule::index => {
                for i_inner in inner.into_inner() {
                    if i_inner.as_rule() == Rule::expression {
                        chain.push(PrimaryChain::Index(walk_expression(i_inner)?));
                    }
                }
            }
            Rule::two_index_slice | Rule::three_index_slice => {
                let slice_source = inner.as_str();
                let mut start = None;
                let mut end = None;
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::expression {
                        if start.is_none() && !slice_source.starts_with("[:") {
                            start = Some(walk_expression(s_inner)?);
                        } else if end.is_none() {
                            end = Some(walk_expression(s_inner)?);
                        }
                    }
                }
                chain.push(PrimaryChain::Slice { start, end });
            }
            Rule::call => {
                let mut args = Vec::new();
                for c_inner in inner.into_inner() {
                    if c_inner.as_rule() == Rule::argument_list {
                        for arg_inner in c_inner.into_inner() {
                            if arg_inner.as_rule() == Rule::argument {
                                let mut spread = false;
                                let mut val = None;
                                for expr_inner in arg_inner.into_inner() {
                                    if expr_inner.as_rule() == Rule::expression {
                                        val = Some(walk_expression(expr_inner)?);
                                    } else if expr_inner.as_rule() == Rule::type_annotation {
                                        val = Some(go_type_arg_expr(walk_type(expr_inner)));
                                    } else if expr_inner.as_rule() == Rule::spread_suffix {
                                        spread = true;
                                    }
                                }
                                if let Some(expr) = val {
                                    args.push(Argument {
                                        value: expr,
                                        name: None,
                                        by_ref: false,
                                        spread,
                                    });
                                }
                            }
                        }
                    }
                }
                chain.push(PrimaryChain::Call(args));
            }
            Rule::type_assertion => {
                for t_inner in inner.into_inner() {
                    if t_inner.as_rule() == Rule::type_annotation {
                        chain.push(PrimaryChain::TypeAssert(walk_type(t_inner)));
                    }
                }
            }
            _ => {}
        }
    }

    if let Some(mut result) = base {
        for item in chain {
            result = match item {
                PrimaryChain::Member(name) => Expression::new(ExprKind::Member {
                    object: Box::new(result),
                    field: name,
                    null_safe: false,
                }),
                PrimaryChain::Index(idx) => Expression::new(ExprKind::Index {
                    object: Box::new(result),
                    index: Box::new(idx),
                    null_safe: false,
                }),
                PrimaryChain::Slice { start, end } => {
                    let start_expr =
                        start.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Int(0))));
                    let mut args = vec![Argument {
                        value: start_expr,
                        name: None,
                        by_ref: false,
                        spread: false,
                    }];
                    if let Some(end_expr) = end {
                        args.push(Argument {
                            value: end_expr,
                            name: None,
                            by_ref: false,
                            spread: false,
                        });
                    }
                    Expression::new(ExprKind::Call {
                        callee: Box::new(Expression::new(ExprKind::Member {
                            object: Box::new(result),
                            field: "slice".to_string(),
                            null_safe: false,
                        })),
                        args,
                        optional: false,
                    })
                }
                PrimaryChain::Call(args) => Expression::new(ExprKind::Call {
                    callee: Box::new(result),
                    args,
                    optional: false,
                }),
                PrimaryChain::TypeAssert(type_name) => go_type_assert_expr(result, type_name),
            };
        }
        Ok(result)
    } else {
        Ok(Expression::new(ExprKind::Lit(Literal::Null)))
    }
}

#[derive(Clone)]
enum PrimaryChain {
    Member(String),
    Index(Expression),
    Slice {
        start: Option<Expression>,
        end: Option<Expression>,
    },
    Call(Vec<Argument>),
    TypeAssert(String),
}

fn walk_operand(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal => return walk_literal(inner),
            Rule::slice_conversion => return walk_slice_conversion(inner),
            Rule::interface_conversion => return walk_type_conversion(inner),
            Rule::ident_name => {
                let name = inner.as_str();
                // Go builtins
                match name {
                    "nil" => return Ok(Expression::new(ExprKind::Lit(Literal::Null))),
                    "true" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(true)))),
                    "false" => return Ok(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
                    _ => return Ok(Expression::new(ExprKind::Ident(name.to_string()))),
                }
            }
            Rule::expression | Rule::if_expression | Rule::range_expression => {
                return walk_expression(inner);
            }
            Rule::composite_literal => return walk_composite_literal(inner),
            Rule::function_literal => return walk_function_literal(inner),
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_slice_conversion(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::from("[]");
    let mut expr = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => type_name.push_str(inner.as_str()),
            Rule::expression => expr = Some(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        type_name,
    }))
}

fn walk_type_conversion(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = None;
    let mut expr = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::type_annotation => type_name = Some(walk_type(inner)),
            Rule::expression => expr = Some(walk_expression(inner)?),
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Cast {
        expr: Box::new(expr.unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))),
        type_name: type_name.unwrap_or_default(),
    }))
}

fn walk_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::numeric_literal => {
                let s = inner.as_str().replace('_', "");
                if s.starts_with("0x") || s.starts_with("0X") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 16) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0b") || s.starts_with("0B") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 2) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.starts_with("0o") || s.starts_with("0O") {
                    if let Ok(n) = i64::from_str_radix(&s[2..], 8) {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                    }
                } else if s.contains('.')
                    || s.contains('e')
                    || s.contains('E')
                    || s.contains('p')
                    || s.contains('P')
                {
                    if let Ok(f) = s.parse::<f64>() {
                        return Ok(Expression::new(ExprKind::Lit(Literal::Float(f))));
                    }
                } else if let Ok(n) = s.parse::<i64>() {
                    return Ok(Expression::new(ExprKind::Lit(Literal::Int(n))));
                }
            }
            Rule::string_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(
                    inner.as_str(),
                )))));
            }
            Rule::bool_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Bool(
                    inner.as_str() == "true",
                ))));
            }
            Rule::nil_literal => {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            }
            Rule::rune_literal => {
                let rune = unquote(inner.as_str());
                let code = rune.chars().next().map(|ch| ch as i64).unwrap_or(0);
                return Ok(Expression::new(ExprKind::Lit(Literal::Int(code))));
            }
            _ => {}
        }
    }
    Ok(Expression::new(ExprKind::Lit(Literal::Null)))
}

fn walk_composite_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut type_name = String::new();
    let mut elements = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::literal_type => {
                type_name = go_literal_type_name(inner);
            }
            Rule::literal_value => {
                for lv_inner in inner.into_inner() {
                    if lv_inner.as_rule() == Rule::element_list {
                        elements = walk_element_list(lv_inner)?;
                    }
                }
            }
            _ => {}
        }
    }

    if type_name.starts_with("map[") {
        // Build a dict/object literal
        let mut props = Vec::new();
        for (key, val) in elements {
            props.push(ObjectProperty::KeyValue { key, value: val });
        }
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Object(props)),
            &type_name,
        ))
    } else if go_is_array_like_type(&type_name) {
        let elem_type = go_array_element_type(&type_name);
        let mut values = Vec::new();
        if let Some(target_len) = go_fixed_array_len(&type_name, elements.len()) {
            if let Some(elem_type) = elem_type.as_deref() {
                values.resize_with(target_len, || go_zero_value_expr(elem_type));
            } else {
                values.resize_with(target_len, Expression::null);
            }
        }
        let mut next_index = 0usize;
        for (key, value) in elements {
            let index = go_composite_literal_index_key(&key).unwrap_or(next_index);
            if index >= values.len() {
                if let Some(elem_type) = elem_type.as_deref() {
                    values.resize_with(index + 1, || go_zero_value_expr(elem_type));
                } else {
                    values.resize_with(index + 1, Expression::null);
                }
            }
            values[index] = go_retype_elided_element(value, elem_type.as_deref());
            next_index = index + 1;
        }
        let arr_elems: Vec<ArrayElement> = values
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Array(arr_elems)),
            &type_name,
        ))
    } else if !type_name.is_empty()
        && elements
            .iter()
            .all(|(key, _)| !matches!(key.kind, ExprKind::Lit(Literal::Null)))
    {
        let mut props = Vec::new();
        for (key, val) in elements {
            let key = match key.kind {
                ExprKind::Ident(name) => Expression::new(ExprKind::Lit(Literal::Str(name))),
                ExprKind::Lit(Literal::Int(n)) => {
                    Expression::new(ExprKind::Lit(Literal::Str(n.to_string())))
                }
                _ => key,
            };
            props.push(ObjectProperty::KeyValue { key, value: val });
        }
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Object(props)),
            &type_name,
        ))
    } else {
        // Untyped composite literal fallback.
        let arr_elems: Vec<ArrayElement> = elements
            .into_iter()
            .map(|(_, v)| ArrayElement {
                key: None,
                value: v,
                spread: false,
                by_ref: false,
            })
            .collect();
        Ok(go_typed_composite_expr(
            Expression::new(ExprKind::Array(arr_elems)),
            &type_name,
        ))
    }
}

/// Apply the composite's element type to an elided element literal. In
/// `[]tagged{{1, 10}}` the inner `{1, 10}` is walked as an untyped composite
/// (a bare `Array`/`Object`); tagging it with the element type lets the
/// normalize pass expand it to the proper struct/slice/map value. Scalars and
/// already-typed elements pass through unchanged.
fn go_retype_elided_element(value: Expression, elem_type: Option<&str>) -> Expression {
    match (elem_type, &value.kind) {
        (Some(elem_type), ExprKind::Array(_) | ExprKind::Object(_)) if !elem_type.is_empty() => {
            go_typed_composite_expr(value, elem_type)
        }
        _ => value,
    }
}

/// Type name of a composite `literal_type`, erasing generic type arguments
/// (`Pair[int]` → `Pair`).
fn go_literal_type_name(pair: Pair<Rule>) -> String {
    if let Some(backing) = go_stdlib_type_binding(pair.as_str()) {
        return backing.to_string();
    }
    let mut inners = pair.clone().into_inner();
    if let Some(first) = inners.next() {
        if first.as_rule() == Rule::ident_name {
            if let Some(second) = inners.next() {
                if second.as_rule() == Rule::type_arguments {
                    return first.as_str().to_string();
                }
            }
        }
    }
    pair.as_str().to_string()
}

fn go_composite_literal_index_key(expr: &Expression) -> Option<usize> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(index)) if *index >= 0 => Some(*index as usize),
        _ => None,
    }
}

fn go_typed_composite_expr(expr: Expression, type_name: &str) -> Expression {
    if type_name.is_empty() {
        expr
    } else {
        Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: type_name.to_string(),
        })
    }
}

fn go_is_array_like_type(type_name: &str) -> bool {
    go_array_head(type_name).is_some()
}

fn go_array_head(type_name: &str) -> Option<(&str, &str)> {
    let trimmed = type_name.trim();
    if !trimmed.starts_with('[') {
        return None;
    }
    let close = trimmed.find(']')?;
    Some((&trimmed[1..close], trimmed[close + 1..].trim()))
}

fn go_array_element_type(type_name: &str) -> Option<String> {
    let (_, tail) = go_array_head(type_name)?;
    (!tail.is_empty()).then(|| tail.to_string())
}

fn go_fixed_array_bounds_exprs(type_name: &str) -> Option<Vec<Expression>> {
    let mut remaining = type_name.trim();
    let mut bounds = Vec::new();

    while let Some((head, tail)) = go_array_head(remaining) {
        let head = head.trim();
        if head.is_empty() || head == "..." {
            return None;
        }
        bounds.push(Expression::int(head.parse::<i64>().ok()?));
        remaining = tail.trim();
    }

    (!bounds.is_empty()).then_some(bounds)
}

fn go_fixed_array_len(type_name: &str, inferred_len: usize) -> Option<usize> {
    let (head, _) = go_array_head(type_name)?;
    let head = head.trim();
    if head.is_empty() {
        None
    } else if head == "..." {
        Some(inferred_len)
    } else {
        head.parse::<usize>().ok()
    }
}

fn go_zero_value_expr(type_name: &str) -> Expression {
    let trimmed = type_name.trim();
    let lower = trimmed.to_ascii_lowercase();

    if let Some(len) = go_fixed_array_len(trimmed, 0) {
        if let Some(elem_type) = go_array_element_type(trimmed) {
            let elements = (0..len)
                .map(|_| ArrayElement {
                    key: None,
                    value: go_zero_value_expr(&elem_type),
                    spread: false,
                    by_ref: false,
                })
                .collect();
            return go_typed_composite_expr(Expression::new(ExprKind::Array(elements)), trimmed);
        }
    }

    if lower.starts_with("[]")
        || lower.starts_with("map[")
        || lower.starts_with("chan ")
        || lower.starts_with('*')
    {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    match lower.as_str() {
        "bool" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        "string" => Expression::new(ExprKind::Lit(Literal::Str(String::new()))),
        "float32" | "float64" => Expression::new(ExprKind::Lit(Literal::Float(0.0))),
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16" | "uint32"
        | "uint64" | "uintptr" | "byte" | "rune" => Expression::new(ExprKind::Lit(Literal::Int(0))),
        _ => go_typed_composite_expr(Expression::new(ExprKind::Object(Vec::new())), trimmed),
    }
}

fn go_zero_value_for_type(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    if let Some(underlying) = env
        .named_types
        .get(type_name)
        .filter(|underlying| underlying.as_str() != type_name)
    {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_zero_value_for_type(underlying, env)),
            type_name: type_name.to_string(),
        });
    }
    go_zero_value_expr(type_name)
}

fn go_map_value_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim();
    if !trimmed.starts_with("map[") {
        return None;
    }

    let mut depth = 0usize;
    for (idx, ch) in trimmed.char_indices() {
        match ch {
            '[' => depth += 1,
            ']' => {
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    let tail = trimmed.get(idx + 1..)?.trim();
                    return (!tail.is_empty()).then(|| tail.to_string());
                }
            }
            _ => {}
        }
    }
    None
}

fn walk_element_list(pair: Pair<Rule>) -> Result<Vec<(Expression, Expression)>, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::keyed_element {
            let parts: Vec<_> = inner.into_inner().collect();
            if parts.len() >= 2 {
                let key = go_keyed_element_key(parts[0].clone())?;
                let value = go_keyed_element_value(parts[1].clone())?;
                elements.push((key, value));
            } else if let Some(value_pair) = parts.into_iter().next() {
                elements.push((
                    Expression::new(ExprKind::Lit(Literal::Null)),
                    go_keyed_element_value(value_pair)?,
                ));
            } else {
                elements.push((
                    Expression::new(ExprKind::Lit(Literal::Null)),
                    Expression::new(ExprKind::Lit(Literal::Null)),
                ));
            }
        }
    }
    Ok(elements)
}

fn go_keyed_element_key(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::ident_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(
            pair.as_str(),
        ))))),
        Rule::bool_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(
            pair.as_str() == "true",
        )))),
        Rule::numeric_literal => {
            let literal = pair.as_str().replace('_', "");
            if let Ok(n) = literal.parse::<i64>() {
                Ok(Expression::new(ExprKind::Lit(Literal::Int(n))))
            } else if let Ok(f) = literal.parse::<f64>() {
                Ok(Expression::new(ExprKind::Lit(Literal::Float(f))))
            } else {
                Ok(Expression::new(ExprKind::Lit(Literal::Null)))
            }
        }
        Rule::expression => walk_expression(pair),
        Rule::element => go_keyed_element_value(pair),
        Rule::literal_value => walk_literal_value_expr(pair),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn go_keyed_element_value(pair: Pair<Rule>) -> Result<Expression, String> {
    match pair.as_rule() {
        Rule::element => {
            let Some(inner) = pair.into_inner().next() else {
                return Ok(Expression::new(ExprKind::Lit(Literal::Null)));
            };
            go_keyed_element_value(inner)
        }
        Rule::expression => walk_expression(pair),
        Rule::literal_value => walk_literal_value_expr(pair),
        Rule::ident_name => Ok(Expression::new(ExprKind::Ident(pair.as_str().to_string()))),
        Rule::string_literal => Ok(Expression::new(ExprKind::Lit(Literal::Str(unquote(
            pair.as_str(),
        ))))),
        Rule::bool_literal => Ok(Expression::new(ExprKind::Lit(Literal::Bool(
            pair.as_str() == "true",
        )))),
        Rule::numeric_literal => go_keyed_element_key(pair),
        _ => Ok(Expression::new(ExprKind::Lit(Literal::Null))),
    }
}

fn walk_literal_value_expr(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut elements = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::element_list {
            elements = walk_element_list(inner)?;
        }
    }

    if elements
        .iter()
        .all(|(key, _)| !matches!(key.kind, ExprKind::Lit(Literal::Null)))
    {
        let mut props = Vec::new();
        for (key, value) in elements {
            let key = match key.kind {
                ExprKind::Ident(name) => Expression::string(&name),
                ExprKind::Lit(Literal::Int(n)) => Expression::string(&n.to_string()),
                _ => key,
            };
            props.push(ObjectProperty::KeyValue { key, value });
        }
        Ok(Expression::new(ExprKind::Object(props)))
    } else {
        Ok(Expression::new(ExprKind::Array(
            elements
                .into_iter()
                .map(|(_, value)| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect(),
        )))
    }
}

fn walk_function_literal(pair: Pair<Rule>) -> Result<Expression, String> {
    let mut params = Vec::new();
    let mut body = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::signature => {
                let sig = walk_signature(inner)?;
                params = sig.params;
            }
            Rule::function_body | Rule::block_statement => {
                body = walk_block(inner)?;
            }
            _ => {}
        }
    }

    Ok(Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    }))
}

fn walk_expression_list(pair: Pair<Rule>) -> Result<Vec<Expression>, String> {
    let mut exprs = Vec::new();
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::expression {
            exprs.push(walk_expression(inner)?);
        }
    }
    Ok(exprs)
}

// ── Helpers ───────────────────────────────────────────────────────────────────────────────

fn first_meaningful(pair: Pair<Rule>) -> Result<Pair<Rule>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() != Rule::EOI {
            return Ok(inner);
        }
    }
    Err("No meaningful child".to_string())
}

fn parse_bin_op(op: &str) -> BinOp {
    match op {
        "+" => BinOp::Add,
        "-" => BinOp::Sub,
        "*" => BinOp::Mul,
        "/" => BinOp::Div,
        "%" => BinOp::Mod,
        "==" => BinOp::Eq,
        "!=" => BinOp::NotEq,
        "<" => BinOp::Lt,
        "<=" => BinOp::LtEq,
        ">" => BinOp::Gt,
        ">=" => BinOp::GtEq,
        "&&" => BinOp::And,
        "||" => BinOp::Or,
        "&" => BinOp::BitAnd,
        "|" => BinOp::BitOr,
        "^" => BinOp::BitXor,
        "<<" => BinOp::Shl,
        ">>" => BinOp::Shr,
        "&^" => BinOp::BitAnd,
        _ => BinOp::Add,
    }
}

fn build_go_binary_expr(op: &str, left: Expression, right: Expression) -> Expression {
    if op == "&^" {
        Expression::new(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(left),
            right: Box::new(Expression::new(ExprKind::Unary {
                op: UnaryOp::BitNot,
                expr: Box::new(right),
            })),
        })
    } else {
        Expression::new(ExprKind::Binary {
            op: parse_bin_op(op),
            left: Box::new(left),
            right: Box::new(right),
        })
    }
}

fn go_reduce_binary_expr(
    operands: &mut Vec<Expression>,
    operators: &mut Vec<String>,
) -> Result<(), String> {
    let Some(op) = operators.pop() else {
        return Ok(());
    };
    let Some(right) = operands.pop() else {
        return Err(format!("missing right operand for Go binary operator {op}"));
    };
    let Some(left) = operands.pop() else {
        return Err(format!("missing left operand for Go binary operator {op}"));
    };
    operands.push(build_go_binary_expr(&op, left, right));
    Ok(())
}

fn go_binary_precedence(op: &str) -> u8 {
    match op {
        "||" => 1,
        "&&" => 2,
        "==" | "!=" | "<" | "<=" | ">" | ">=" => 3,
        "+" | "-" | "|" | "^" => 4,
        "*" | "/" | "%" | "<<" | ">>" | "&" | "&^" => 5,
        _ => 0,
    }
}

fn go_type_arg_expr(type_name: String) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::null()),
        type_name,
    })
}

fn go_type_assert_expr(expr: Expression, type_name: String) -> Expression {
    go_builtin_call("__go_type_assert", vec![expr, go_type_arg_expr(type_name)])
}

fn go_extract_type_assert_expr(expr: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if !matches!(callee.kind, ExprKind::Ident(ref name) if name == "__go_type_assert")
        || args.len() != 2
    {
        return None;
    }
    Some((
        args[0].value.clone(),
        go_type_name_from_expr(&args[1].value)?,
    ))
}

fn go_type_assert_value_expr(expr: Expression, type_name: &str) -> Expression {
    if !matches!(
        type_name.trim(),
        "int"
            | "int8"
            | "int16"
            | "int32"
            | "int64"
            | "uint"
            | "uint8"
            | "uint16"
            | "uint32"
            | "uint64"
            | "uintptr"
            | "byte"
            | "rune"
            | "float32"
            | "float64"
            | "string"
            | "bool"
    ) {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: type_name.to_string(),
        });
    }

    let cond = go_build_is_type(expr.clone(), type_name);
    let then_expr = Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then_expr),
        else_: Box::new(go_zero_value_expr(type_name)),
    })
}

fn go_type_switch_case_cond(expr: Expression, case_types: &[String]) -> Expression {
    let mut iter = case_types.iter();
    let first = iter
        .next()
        .map(|type_name| go_build_is_type(expr.clone(), type_name))
        .unwrap_or_else(|| Expression::bool(false));
    iter.fold(first, |acc, type_name| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(acc),
            right: Box::new(go_build_is_type(expr.clone(), type_name)),
        })
    })
}

fn go_build_is_type(expr: Expression, type_name: &str) -> Expression {
    let typeof_tag = match type_name.trim() {
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16" | "uint32"
        | "uint64" | "uintptr" | "byte" | "rune" | "float32" | "float64" => Some("number"),
        "string" => Some("string"),
        "bool" => Some("boolean"),
        _ => None,
    };

    if let Some(tag) = typeof_tag {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::string(tag)),
        });
    }

    // Map Go composite types to the canonical IsType categories the shared
    // compiler recognizes (array→isArray, function, map→object-kind).
    let trimmed = type_name.trim();
    let canon = if trimmed.starts_with("[]")
        || (trimmed.starts_with('[') && trimmed.contains(']'))
    {
        "array"
    } else if trimmed.starts_with("func") {
        "function"
    } else if trimmed.starts_with("map[") {
        "map"
    } else {
        trimmed
    };

    Expression::new(ExprKind::IsType {
        expr: Box::new(expr),
        type_name: canon.to_string(),
    })
}

fn go_type_switch_case_body(
    mut body: Vec<Statement>,
    binding_name: Option<&str>,
    expr: Expression,
    case_type: &str,
) -> Vec<Statement> {
    if let Some(name) = binding_name {
        body.insert(
            0,
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name.to_string()),
                    init: Some(Expression::new(ExprKind::Cast {
                        expr: Box::new(expr),
                        type_name: case_type.to_string(),
                    })),
                    type_hint: Some(case_type.to_string()),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }),
        );
    }
    body
}

fn go_wrap_spawn_expr(expr: Expression) -> Expression {
    Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Expr(expr))]),
        is_async: false,
        captures: Vec::new(),
    })
}

fn go_type_name_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Cast { expr, type_name } if matches!(expr.kind, ExprKind::Lit(Literal::Null)) => {
            Some(type_name.clone())
        }
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { .. } => go_expr_call_name(expr),
        _ => None,
    }
}

fn go_is_slice_type(type_name: &str) -> bool {
    type_name.trim_start().starts_with("[]")
}

fn go_is_map_type(type_name: &str) -> bool {
    type_name.trim_start().starts_with("map[")
}

fn go_is_channel_type(type_name: &str) -> bool {
    let trimmed = type_name.trim_start();
    trimmed.starts_with("chan") || trimmed.starts_with("<-chan")
}

fn go_array_make_expr(len_expr: Expression, init_expr: Expression) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(len_expr),
            Argument::positional(init_expr),
        ],
        optional: false,
    })
}

fn go_make_slice_capacity_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("make") {
        return None;
    }
    let type_name = args
        .first()
        .and_then(|arg| go_type_name_from_expr(&arg.value))?;
    if !go_is_slice_type(&type_name) {
        return None;
    }
    let cap_arg = args.get(2).or_else(|| args.get(1))?;
    Some(normalize_go_expr(&cap_arg.value, env, signatures, state))
}

fn go_bound_slice_capacity_expr(expr: &Expression, env: &GoNormalizeEnv) -> Option<Expression> {
    if let Some(cap) = go_expr_capacity_hint(expr, env) {
        return Some(cap);
    }

    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            if go_expr_call_name(callee).as_deref() != Some("make") {
                return None;
            }
            let type_name = args
                .first()
                .and_then(|arg| go_type_name_from_expr(&arg.value))?;
            if !go_is_slice_type(&type_name) {
                return None;
            }
            Some(args.get(2).or_else(|| args.get(1))?.value.clone())
        }
        ExprKind::Cast { expr, type_name } if go_is_slice_type(type_name) => {
            go_bound_slice_capacity_expr(expr, env)
        }
        ExprKind::Array(elements) => Some(Expression::int(elements.len() as i64)),
        _ => None,
    }
}

fn go_expr_capacity_hint(expr: &Expression, env: &GoNormalizeEnv) -> Option<Expression> {
    if let Some(view) = go_expr_slice_view(expr, env) {
        let base_cap = go_expr_capacity_hint(&view.base, env)?;
        return Some(Expression::new(ExprKind::Binary {
            op: BinOp::Sub,
            left: Box::new(base_cap),
            right: Box::new(view.start),
        }));
    }
    match &expr.kind {
        ExprKind::Ident(name) => env.slice_caps.get(name).cloned().or_else(|| {
            env.fixed_arrays
                .get(name)
                .and_then(|type_name| go_fixed_array_len(type_name, 0))
                .map(|len| Expression::int(len as i64))
        }),
        ExprKind::Array(elements) => Some(Expression::int(elements.len() as i64)),
        _ => None,
    }
}

fn go_binding_name(pattern: &BindingPattern) -> Option<String> {
    match pattern {
        BindingPattern::Ident(name) => Some(name.clone()),
        _ => go_single_named_binding_pattern(pattern).and_then(|pattern| match pattern {
            BindingPattern::Ident(name) => Some(name),
            _ => None,
        }),
    }
}

fn to_span(pair: &Pair<Rule>) -> Span {
    let s = pair.as_span();
    let (sl, sc) = s.start_pos().line_col();
    let (el, ec) = s.end_pos().line_col();
    Span {
        start_line: sl as u32,
        start_col: sc as u32,
        end_line: el as u32,
        end_col: ec as u32,
    }
}
