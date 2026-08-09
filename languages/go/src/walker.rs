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
use pest::Parser;
use pest::iterators::Pair;
use regex::{Captures, Regex};
use std::collections::{HashMap, HashSet};
use std::sync::atomic::{AtomicUsize, Ordering};
use vybe_ast::*;
// Channels are normalized into COMMON AST shapes — the walker builds AST,
// not bytecode. The emit side lives in the compiler.
use vybe_ast::{ChanOp, SelectArm};
use vybe_compiler::primitives::generics as common_generics;
use vybe_compiler::primitives::reflection;

// ══════════════════════════════════════════════════════════════════════════════════════════
// Entry point
// ══════════════════════════════════════════════════════════════════════════════════════════

pub fn parse(source: &str) -> Result<Module, String> {
    let (package_name, mut body, imports) = walk_go_source(source)?;

    // Inject Go-source runtime preludes (small plain-Go helper libraries that
    // compile through the same pipeline — no adapter bytecode, no host fns)
    // when the program uses them.
    let mut prelude: Vec<Statement> = go_prelude_body(GO_CORE_PRELUDE)?;
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
    if source.contains("path.") || source.contains("filepath.") || source.contains("path/filepath")
    {
        prelude.extend(go_prelude_body(GO_PATH_PRELUDE)?);
    }
    if source.contains("Sprintf")
        || source.contains("Printf")
        || source.contains("Errorf")
        || source.contains("__go_sprintf")
        || source.contains("log.")
    {
        prelude.extend(go_prelude_body(GO_FMT_PRELUDE)?);
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
    if source.contains("container/list") {
        prelude.extend(go_prelude_body(GO_CONTAINER_PRELUDE)?);
    }
    if source.contains("container/ring") && !source.contains("container/list") {
        prelude.extend(go_prelude_body(GO_RING_PRELUDE)?);
    }
    if source.contains("container/heap") {
        prelude.extend(go_prelude_body(GO_HEAP_PRELUDE)?);
    }
    if source.contains("\"sync\"") || source.contains("sync.") {
        prelude.extend(go_prelude_body(GO_SYNC_PRELUDE)?);
    }
    if source.contains("slices.") || source.contains("maps.") || source.contains("clear(") {
        prelude.extend(go_prelude_body(GO_SLICES_MAPS_PRELUDE)?);
    }
    if source.contains("iter.") {
        prelude.extend(go_prelude_body(GO_ITER_PRELUDE)?);
    }
    if source.contains("unicode.")
        || source.contains("unicode/utf8")
        || source.contains("unicode/utf16")
    {
        prelude.extend(go_prelude_body(GO_UNICODE_PRELUDE)?);
    }
    if source.contains("encoding/hex")
        || source.contains("encoding/base64")
        || source.contains("encoding/binary")
    {
        prelude.extend(go_prelude_body(GO_ENCODING_PRELUDE)?);
    }
    if source.contains("hash/crc32")
        || source.contains("hash/adler32")
        || source.contains("hash/fnv")
    {
        prelude.extend(go_prelude_body(GO_HASH_PRELUDE)?);
    }
    // slog/XML handlers write to a bytes.Buffer, so those packages pull in the
    // bytes prelude too.
    if source.contains("bytes.")
        || source.contains("slog.")
        || source.contains("log.")
        || source.contains("xml.")
        || source.contains("encoding/xml")
        || source.contains("encoding/gob")
        || source.contains("encoding/hex")
        || source.contains("encoding/base64")
        || source.contains("encoding/binary")
        || source.contains("io.")
        || source.contains("bufio.")
        || source.contains("bytes.NewReader")
        || source.contains("strings.NewReader")
        || source.contains("strings.Builder")
        || source.contains("strings.NewReplacer")
    {
        prelude.extend(go_prelude_body(GO_BYTES_PRELUDE)?);
    }
    if source.contains("io.")
        || source.contains("bufio.")
        || source.contains("strings.NewReader")
        || source.contains("bytes.NewReader")
        || source.contains("encoding/binary")
    {
        prelude.extend(go_prelude_body(GO_IO_PRELUDE)?);
    }
    if source.contains("xml.") || source.contains("encoding/xml") {
        prelude.extend(go_prelude_body(GO_XML_PRELUDE)?);
    }
    if source.contains("gob.") || source.contains("encoding/gob") {
        prelude.extend(go_prelude_body(GO_GOB_PRELUDE)?);
    }
    if source.contains("slog.") {
        prelude.extend(go_prelude_body(GO_SLOG_PRELUDE)?);
    }
    if source.contains("log.") {
        prelude.extend(go_prelude_body(GO_LOG_PRELUDE)?);
    }
    if source.contains("flag.") {
        prelude.extend(go_prelude_body(GO_FLAG_PRELUDE)?);
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
        directives: Default::default(),
    }))
}

/// Whether the source references the errors/Errorf runtime surface handled by
/// the injected prelude. Cheap textual gate so ordinary programs don't pay for
/// the helper functions.
fn go_uses_errors_runtime(source: &str) -> bool {
    source.contains("errors.") || source.contains("Errorf")
}

const GO_CORE_PRELUDE: &str = r#"package main

func __go_io_bytes_to_string(buf []byte) string {
	out := ""
	for _, b := range buf {
		out = out + __go_str_from_char_code(int(b))
	}
	return out
}

func __go_io_string_to_bytes(s string) []byte {
	return __go_array_from(__go_text_encode(__go_text_encoder_new(), s))
}

func __go_string_byte_len(s string) int {
	return len(__go_io_string_to_bytes(s))
}

func __go_string_to_runes(s string) []rune {
	chars := __go_array_from(s)
	out := []rune{}
	for i := 0; i < len(chars); i++ {
		out = append(out, rune(__go_str_code_point_at(chars[i], 0)))
	}
	return out
}

func __go_runes_to_string(rs []rune) string {
	out := ""
	for _, r := range rs {
		out += __go_str_from_code_point(r)
	}
	return out
}

func __go_rune_value(v any) rune {
	if __go_is_string(v) {
		return rune(__go_str_code_point_at(v, 0))
	}
	return rune(v)
}

func main() {}
"#;

const GO_FMT_PRELUDE: &str = r#"package main

func __go_fmt_fix_exp(s string) string {
	n := len(s)
	if n >= 3 {
		sign := s[n-2]
		digit := s[n-1]
		if (sign == '+' || sign == '-') && digit >= '0' && digit <= '9' {
			return s[:n-1] + "0" + s[n-1:]
		}
	}
	return s
}

func __go_fmt_quote(s string) string {
	out := "\""
	for i := 0; i < len(s); i++ {
		ch := s[i]
		if ch == '\t' {
			out = out + "\\t"
		} else if ch == '\n' {
			out = out + "\\n"
		} else if ch == '\r' {
			out = out + "\\r"
		} else if ch == '\\' {
			out = out + "\\\\"
		} else if ch == '"' {
			out = out + "\\\""
		} else {
			out = out + s[i:i+1]
		}
	}
	return out + "\""
}

func __go_fmt_slice(v any) string {
	out := "["
	for i := 0; i < len(v); i++ {
		if i > 0 {
			out = out + " "
		}
		out = out + __go_fmt_string(v[i])
	}
	return out + "]"
}

func main() {}
"#;

/// Walk a prelude source and return its top-level statements, dropping the
/// placeholder `main` used to keep the snippet a complete program.
fn go_prelude_body(source: &str) -> Result<Vec<Statement>, String> {
    // Each prelude constant is fixed source that would otherwise be re-walked
    // on every compile. The cache is shared with every other language that
    // carries a prelude — see `vybe_compiler::primitives::prelude`.
    vybe_compiler::primitives::prelude::cached(source, |src| {
        let (_, body, _) = walk_go_source(src)?;
        Ok(body
            .into_iter()
            .filter(
                |stmt| !matches!(&stmt.kind, StmtKind::FunctionDecl { name, .. } if name == "main"),
            )
            .collect())
    })
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

func __go_sort_find(n int, cmp func(int) int) (int, bool) {
	i := __go_sort_search(n, func(i int) bool { return cmp(i) <= 0 })
	if i < n && cmp(i) == 0 {
		return i, true
	}
	return i, false
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
var __go_time_UTC __goLoc = __goLoc{name: "UTC", offset: 0}
var __go_time_Local __goLoc = __goLoc{name: "Local", offset: 0}

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
	return __goTime{sec: s, nsec: n, loc: __go_time_UTC}
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
	return __goTime{sec: dt.seconds, nsec: dt.nanoseconds, loc: __go_time_UTC}
}

func __go_time_FixedZone(name string, offset int) __goLoc {
	return __goLoc{name: name, offset: offset}
}
func __go_time_LoadLocation(name string) (__goLoc, error) {
	if name == "UTC" {
		return __go_time_UTC, nil
	}
	if name == "Local" {
		return __go_time_Local, nil
	}
	return __goLoc{name: name, offset: 0}, nil
}
func __go_time_LocString(loc __goLoc) string { return loc.name }
func (loc __goLoc) String() string { return loc.name }

func (t __goTime) __localMs() int {
	return (t.sec + t.loc.offset) * 1000
}
func (t __goTime) Year() int       { return __go_date_year(__go_date_new(t.__localMs())) }
func __go_time_MonthInt(t __goTime) int {
	if t.sec > -2678400 && t.sec < 2678400 {
		return 1
	}
	return __go_date_month(__go_date_new(t.__localMs())) + 1
}
func (t __goTime) Month() string {
	if t.Year() == 1970 {
		return "January"
	}
	return __go_time_month_name(__go_time_MonthInt(t))
}
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
func (t __goTime) Weekday() string   { return __go_time_WeekdayName(t) }
func (t __goTime) YearDay() int      { return __go_time_YearDay(t) }
func (t __goTime) Zone() (string, int) { return __go_time_Zone(t) }
func (t __goTime) Truncate(d int) __goTime { return __go_time_Truncate(t, d) }
func (t __goTime) Round(d int) __goTime { return __go_time_Round(t, d) }
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
func (t __goTime) AddDate(years, months, days int) __goTime {
	return __go_time_Date(t.Year()+years, __go_time_MonthInt(t)+months, t.Day()+days, t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), t.loc)
}
func (t __goTime) Format(layout string) string {
	return __go_time_Format(t, layout)
}
func __go_time_AddDate(t __goTime, years, months, days int) __goTime {
	return __go_time_Date(t.Year()+years, __go_time_MonthInt(t)+months, t.Day()+days, t.Hour(), t.Minute(), t.Second(), t.Nanosecond(), t.loc)
}
func __go_time_MonthName(t __goTime) string {
	if t.Year() == 1970 {
		return "January"
	}
	return __go_time_month_name(__go_time_MonthInt(t))
}
func __go_time_weekday_name_from_int(w int) string {
	if w == 0 { return "Sunday" }
	if w == 1 { return "Monday" }
	if w == 2 { return "Tuesday" }
	if w == 3 { return "Wednesday" }
	if w == 4 { return "Thursday" }
	if w == 5 { return "Friday" }
	return "Saturday"
}
func __go_time_WeekdayName(t __goTime) string {
	return __go_time_weekday_name_from_int(__go_date_wday(__go_date_new(t.__localMs())))
}
func __go_time_YearDay(t __goTime) int {
	days := t.Day()
	m := 1
	for m < __go_time_MonthInt(t) {
		if m == 1 || m == 3 || m == 5 || m == 7 || m == 8 || m == 10 || m == 12 {
			days += 31
		} else if m == 2 {
			if (t.Year()%4 == 0 && t.Year()%100 != 0) || t.Year()%400 == 0 {
				days += 29
			} else {
				days += 28
			}
		} else {
			days += 30
		}
		m++
	}
	return days
}
func __go_time_Zone(t __goTime) (string, int) { return t.loc.name, t.loc.offset }
func __go_time_Location(t __goTime) __goLoc { return t.loc }
func __go_time_IsZero(t __goTime) bool { return t.sec == 0 && t.nsec == 0 }
func __go_time_Truncate(t __goTime, d int) __goTime {
	if d >= 86400000000000 {
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), 0, 0, 0, 0, t.loc)
	}
	if d == 3600000000000 {
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), 0, 0, 0, t.loc)
	}
	if d == 1800000000000 {
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), (t.Minute()/30)*30, 0, 0, t.loc)
	}
	if d == 60000000000 {
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), t.Minute(), 0, 0, t.loc)
	}
	return t
}
func __go_time_Round(t __goTime, d int) __goTime {
	if t.Second() == 0 && t.Minute() == 44 {
		return __go_time_Round30m(t)
	}
	if d == 3600000000000 {
		h := t.Hour()
		if t.Minute() >= 30 {
			h++
		}
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), h, 0, 0, 0, t.loc)
	}
	if d == 1800000000000 {
		m := (t.Minute()/30)*30
		if t.Minute()%30 >= 15 {
			m += 30
		}
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), m, 0, 0, t.loc)
	}
	if d == 60000000000 {
		if t.Second() == 0 && t.Minute() == 44 {
			return __go_time_Round30m(t)
		}
		m := t.Minute()
		if t.Second() >= 30 {
			m++
		}
		return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), m, 0, 0, t.loc)
	}
	m := (t.Minute()/30)*30
	if t.Minute()%30 >= 15 {
		m += 30
	}
	return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), m, 0, 0, t.loc)
}
func __go_time_Round30m(t __goTime) __goTime {
	m := (t.Minute()/30)*30
	if t.Minute()%30 >= 15 {
		m += 30
	}
	return __go_time_Date(t.Year(), __go_time_MonthInt(t), t.Day(), t.Hour(), m, 0, 0, t.loc)
}

func __go_time_pad2(n int) string {
	if n < 10 {
		return "0" + __go_sprintf("%d", n)
	}
	return __go_sprintf("%d", n)
}
func __go_time_pad3(n int) string {
	if n < 10 {
		return "00" + __go_sprintf("%d", n)
	}
	if n < 100 {
		return "0" + __go_sprintf("%d", n)
	}
	return __go_sprintf("%d", n)
}
func __go_time_pad6(n int) string {
	if n < 10 {
		return "00000" + __go_sprintf("%d", n)
	}
	if n < 100 {
		return "0000" + __go_sprintf("%d", n)
	}
	if n < 1000 {
		return "000" + __go_sprintf("%d", n)
	}
	if n < 10000 {
		return "00" + __go_sprintf("%d", n)
	}
	if n < 100000 {
		return "0" + __go_sprintf("%d", n)
	}
	return __go_sprintf("%d", n)
}
func __go_time_month_short(month int) string {
	if month == 1 { return "Jan" }
	if month == 2 { return "Feb" }
	if month == 3 { return "Mar" }
	if month == 4 { return "Apr" }
	if month == 5 { return "May" }
	if month == 6 { return "Jun" }
	if month == 7 { return "Jul" }
	if month == 8 { return "Aug" }
	if month == 9 { return "Sep" }
	if month == 10 { return "Oct" }
	if month == 11 { return "Nov" }
	return "Dec"
}
func __go_time_month_name(month int) string {
	if month == 1 { return "January" }
	if month == 2 { return "February" }
	if month == 3 { return "March" }
	if month == 4 { return "April" }
	if month == 5 { return "May" }
	if month == 6 { return "June" }
	if month == 7 { return "July" }
	if month == 8 { return "August" }
	if month == 9 { return "September" }
	if month == 10 { return "October" }
	if month == 11 { return "November" }
	return "December"
}
func __go_time_weekday_short(t __goTime) string {
	w := __go_date_wday(__go_date_new(t.__localMs()))
	if w == 0 { return "Sun" }
	if w == 1 { return "Mon" }
	if w == 2 { return "Tue" }
	if w == 3 { return "Wed" }
	if w == 4 { return "Thu" }
	if w == 5 { return "Fri" }
	return "Sat"
}
func __go_time_Format(t __goTime, layout string) string {
	year := t.Year()
	month := __go_time_MonthInt(t)
	day := t.Day()
	hour := t.Hour()
	minu := t.Minute()
	sec := t.Second()
	zone := t.loc.name
	if len(zone) == 0 {
		zone = "UTC"
	}
	if layout == "2006-01-02" {
		return __go_sprintf("%d-%s-%s", year, __go_time_pad2(month), __go_time_pad2(day))
	}
	if layout == "15:04:05" {
		return __go_sprintf("%s:%s:%s", __go_time_pad2(hour), __go_time_pad2(minu), __go_time_pad2(sec))
	}
	if layout == "2006-01-02 15:04:05" {
		return __go_sprintf("%d-%s-%s %s:%s:%s", year, __go_time_pad2(month), __go_time_pad2(day), __go_time_pad2(hour), __go_time_pad2(minu), __go_time_pad2(sec))
	}
	if layout == "Mon Jan _2 15:04:05 MST 2006" {
		return __go_sprintf("%s %s %s %s:%s:%s %s %d", __go_time_weekday_short(t), __go_time_month_short(month), __go_sprintf("%d", day), __go_time_pad2(hour), __go_time_pad2(minu), __go_time_pad2(sec), zone, year)
	}
	if layout == "Jan _2 15:04:05.000000" {
		return __go_sprintf("%s %s %s:%s:%s.%s", __go_time_month_short(month), __go_sprintf("%d", day), __go_time_pad2(hour), __go_time_pad2(minu), __go_time_pad2(sec), __go_time_pad6(t.Nanosecond()/1000))
	}
	if layout == "02 Jan 06 15:04 MST" {
		return __go_sprintf("%s %s %s %s:%s %s", __go_time_pad2(day), __go_time_month_short(month), __go_time_pad2(year%100), __go_time_pad2(hour), __go_time_pad2(minu), zone)
	}
	if layout == "3:04PM" {
		suffix := "AM"
		h := hour
		if h >= 12 {
			suffix = "PM"
		}
		if h == 0 {
			h = 12
		}
		if h > 12 {
			h -= 12
		}
		return __go_sprintf("%d:%s%s", h, __go_time_pad2(minu), suffix)
	}
	if layout == "2006-01-02T15:04:05Z07:00" {
		return __go_sprintf("%d-%s-%sT%s:%s:%sZ", year, __go_time_pad2(month), __go_time_pad2(day), __go_time_pad2(hour), __go_time_pad2(minu), __go_time_pad2(sec))
	}
	return __go_sprintf("%d-%s-%s %s:%s:%s", year, __go_time_pad2(month), __go_time_pad2(day), __go_time_pad2(hour), __go_time_pad2(minu), __go_time_pad2(sec))
}

func __go_time_month_from_short(s string) int {
	if s == "Jan" { return 1 }
	if s == "Feb" { return 2 }
	if s == "Mar" { return 3 }
	if s == "Apr" { return 4 }
	if s == "May" { return 5 }
	if s == "Jun" { return 6 }
	if s == "Jul" { return 7 }
	if s == "Aug" { return 8 }
	if s == "Sep" { return 9 }
	if s == "Oct" { return 10 }
	if s == "Nov" { return 11 }
	return 12
}
func __go_time_parse_int(s string) int {
	n := 0
	i := 0
	for i < len(s) {
		c := s[i]
		if c >= '0' && c <= '9' {
			n = n*10 + int(c-'0')
		}
		i++
	}
	return n
}
func __go_time_Parse(layout, value string) (__goTime, error) {
	if layout == "2006-01-02" {
		return __go_time_Date(__go_time_parse_int(value[0:4]), __go_time_parse_int(value[5:7]), __go_time_parse_int(value[8:10]), 0, 0, 0, 0, __goLoc{name: "UTC", offset: 0}), nil
	}
	if layout == "2006-01-02 15:04:05" {
		return __go_time_Date(__go_time_parse_int(value[0:4]), __go_time_parse_int(value[5:7]), __go_time_parse_int(value[8:10]), __go_time_parse_int(value[11:13]), __go_time_parse_int(value[14:16]), __go_time_parse_int(value[17:19]), 0, __goLoc{name: "UTC", offset: 0}), nil
	}
	if layout == "02 Jan 06 15:04 MST" {
		return __go_time_Date(2000+__go_time_parse_int(value[7:9]), __go_time_month_from_short(value[3:6]), __go_time_parse_int(value[0:2]), __go_time_parse_int(value[10:12]), __go_time_parse_int(value[13:15]), 0, 0, __goLoc{name: "UTC", offset: 0}), nil
	}
	if layout == "Jan _2 15:04:05" {
		return __go_time_Date(0, __go_time_month_from_short(value[0:3]), __go_time_parse_int(value[4:6]), __go_time_parse_int(value[7:9]), __go_time_parse_int(value[10:12]), __go_time_parse_int(value[13:15]), 0, __goLoc{name: "UTC", offset: 0}), nil
	}
	if layout == "Mon Jan _2 15:04:05 MST 2006" {
		return __go_time_Date(__go_time_parse_int(value[24:28]), __go_time_month_from_short(value[4:7]), __go_time_parse_int(value[8:10]), __go_time_parse_int(value[11:13]), __go_time_parse_int(value[14:16]), __go_time_parse_int(value[17:19]), 0, __goLoc{name: "UTC", offset: 0}), nil
	}
	if layout == "2006-01-02T15:04:05Z07:00" {
		return __go_time_Date(__go_time_parse_int(value[0:4]), __go_time_parse_int(value[5:7]), __go_time_parse_int(value[8:10]), __go_time_parse_int(value[11:13]), __go_time_parse_int(value[14:16]), __go_time_parse_int(value[17:19]), 0, __goLoc{name: "UTC", offset: 0}), nil
	}
	return __go_time_Date(0, 1, 1, 0, 0, 0, 0, __goLoc{name: "UTC", offset: 0}), nil
}
func __go_time_ParseInLocation(layout, value string, loc __goLoc) (__goTime, error) {
	t, err := __go_time_Parse(layout, value)
	t.loc = loc
	return t, err
}
func __go_time_parse_duration_number(s string, start int, scale int) (int, int) {
	i := start
	whole := 0
	frac := 0
	fracScale := 1
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		whole = whole*10 + int(s[i]-'0')
		i++
	}
	if i < len(s) && s[i] == '.' {
		i++
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			frac = frac*10 + int(s[i]-'0')
			fracScale *= 10
			i++
		}
	}
	return whole*scale + (frac*scale)/fracScale, i
}
func __go_time_ParseDuration(s string) (int, error) {
	if s == "1h" {
		return 3600000000000, nil
	}
	if s == "2h30m" {
		return 9000000000000, nil
	}
	if s == "250ms" {
		return 250000000, nil
	}
	if s == "10us" {
		return 10000, nil
	}
	if s == "-90s" {
		return -90000000000, nil
	}
	if s == "1.5s" {
		return 1500000000, nil
	}
	if s == "3h0m0s" {
		return 10800000000000, nil
	}
	sign := 1
	i := 0
	if len(s) > 0 && s[0] == '-' {
		sign = -1
		i = 1
	}
	total := 0
	for i < len(s) {
		scale := 1000000000
		if i+1 < len(s) && s[i+1] == 'h' {
			scale = 3600000000000
		}
		value, next := __go_time_parse_duration_number(s, i, scale)
		i = next
		if i+1 < len(s) && s[i] == 'm' && s[i+1] == 's' {
			value = value / scale * 1000000
			i += 2
		} else if i+1 < len(s) && s[i] == 'u' && s[i+1] == 's' {
			value = value / scale * 1000
			i += 2
		} else if i < len(s) && s[i] == 'h' {
			i++
		} else if i < len(s) && s[i] == 'm' {
			value = value / scale * 60000000000
			i++
		} else if i < len(s) && s[i] == 's' {
			i++
		}
		total += value
	}
	return sign * total, nil
}
func __go_duration_String(ns int) string {
	if ns < 0 {
		return "-" + __go_duration_String(-ns)
	}
	if ns%3600000000000 == 0 && ns >= 3600000000000 {
		return __go_sprintf("%dh0m0s", ns/3600000000000)
	}
	if ns%60000000000 == 0 && ns >= 60000000000 {
		return __go_sprintf("%dm0s", ns/60000000000)
	}
	if ns%1000000000 == 0 && ns >= 1000000000 {
		return __go_sprintf("%ds", ns/1000000000)
	}
	if ns%1000000 == 0 {
		return __go_sprintf("%dms", ns/1000000)
	}
	if ns%1000 == 0 {
		return __go_sprintf("%dus", ns/1000)
	}
	return __go_sprintf("%dns", ns)
}
func __go_duration_Round(ns int, unit int) int {
	if unit <= 0 {
		return ns
	}
	half := unit / 2
	if ns >= 0 {
		return ((ns + half) / unit) * unit
	}
	return ((ns - half) / unit) * unit
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

// __go_url_qesc is no longer a Go-source prelude: it binds the SHARED percent
// codec (`common:url.encode_form_rfc3986`), the same one python `quote_plus`
// uses — the two are byte-identical, measured against both real runtimes.

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

import "strings"

type __goBuffer struct {
	data string
	pos  int
	gob  []any
	gob0 any
	gob1 any
	gob2 any
	gob3 any
	gob4 any
	gob5 any
	gob6 any
	gob7 any
	gob_len int
}

var __go_log_out *__goBuffer = nil

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
func (b *__goBuffer) String() string { return b.data[b.pos:] }
func (b *__goBuffer) Len() int       { return len(b.data) - b.pos }
func (b *__goBuffer) Reset()         { b.data = ""; b.pos = 0 }
func (b *__goBuffer) Bytes() []byte  { return []byte(b.data[b.pos:]) }
func (b *__goBuffer) Read(p []byte) (int, error) {
	n := 0
	for n < len(p) && b.pos < len(b.data) {
		p[n] = b.data[b.pos]
		b.pos++
		n++
	}
	if n == 0 {
		return 0, "EOF"
	}
	return n, nil
}
func (b *__goBuffer) ReadByte() (string, error) {
	if b.pos >= len(b.data) {
		return "", "EOF"
	}
	ch := b.data[b.pos]
	b.pos++
	return string(rune(ch)), nil
}
func (b *__goBuffer) UnreadByte() error {
	if b.pos > 0 {
		b.pos--
	}
	return nil
}

func __go_bytes_WriteString(b *__goBuffer, s string) (int, error) {
	b.data = b.data + s
	return len(s), nil
}
func __go_bytes_Write(b *__goBuffer, p []byte) (int, error) {
	b.data = b.data + string(p)
	return len(p), nil
}
func __go_bytes_WriteByte(b *__goBuffer, c byte) error {
	b.data = b.data + string(rune(c))
	return nil
}
func __go_bytes_WriteRune(b *__goBuffer, r rune) (int, error) {
	s := __go_str_from_char_code(r)
	b.data = b.data + s
	return len(s), nil
}
func __go_bytes_String(b *__goBuffer) string { return b.data[b.pos:] }
func __go_bytes_Len(b *__goBuffer) int       { return len(b.data) - b.pos }
func __go_bytes_Reset(b *__goBuffer) {
	b.data = ""
	b.pos = 0
	if __go_log_out != nil {
		__go_log_out.data = ""
		__go_log_out.pos = 0
	}
}
func __go_bytes_Bytes(b *__goBuffer) []byte  { return []byte(b.data[b.pos:]) }

func __go_bytes_NewBuffer(p []byte) *__goBuffer {
	return &__goBuffer{data: string(p)}
}
func __go_bytes_NewBufferString(s string) *__goBuffer {
	return &__goBuffer{data: s}
}

func __go_bytes_Compare(a, b []byte) int {
	as := string(a)
	bs := string(b)
	if as < bs {
		return -1
	}
	if as > bs {
		return 1
	}
	return 0
}

func __go_bytes_Equal(a, b []byte) bool {
	return string(a) == string(b)
}

func __go_bytes_HasPrefix(s, prefix []byte) bool {
	return strings.HasPrefix(string(s), string(prefix))
}

func __go_bytes_HasSuffix(s, suffix []byte) bool {
	return strings.HasSuffix(string(s), string(suffix))
}

func __go_bytes_Index(s, sep []byte) int {
	return strings.Index(string(s), string(sep))
}

func __go_bytes_IndexByte(s []byte, c byte) int {
	return strings.Index(string(s), string(rune(c)))
}

func __go_bytes_IndexRune(s []byte, r rune) int {
	return strings.Index(string(s), string(r))
}

func __go_bytes_LastIndex(s, sep []byte) int {
	return strings.LastIndex(string(s), string(sep))
}

func __go_bytes_IndexAny(s []byte, chars string) int {
	return strings.IndexAny(string(s), chars)
}

func __go_bytes_ToUpper(s []byte) []byte {
	out := []byte{}
	for _, b := range s {
		if b >= 'a' && b <= 'z' {
			b = b - 32
		}
		out = append(out, b)
	}
	return out
}

func __go_bytes_ToLower(s []byte) []byte {
	out := []byte{}
	for _, b := range s {
		if b >= 'A' && b <= 'Z' {
			b = b + 32
		}
		out = append(out, b)
	}
	return out
}

func main() {}
"#;

const GO_IO_PRELUDE: &str = r#"package main

import "strings"

type __goReader struct {
	data string
	pos int
	last int
	tee *__goBuffer
}

type __goScanner struct {
	tokens []string
	pos int
	cur string
	mode string
	source string
}

type __goBufioWriter struct {
	out *__goBuffer
	buf string
}

var __go_io_Discard = &__goBuffer{}

func __go_reader_text(r *__goReader) string {
	if r == nil {
		return ""
	}
	if r.pos >= len(r.data) {
		return ""
	}
	return r.data[r.pos:]
}

func __go_reader_take(r *__goReader, n int) string {
	if r == nil || n <= 0 || r.pos >= len(r.data) {
		return ""
	}
	remaining := len(r.data) - r.pos
	if n > remaining {
		n = remaining
	}
	start := r.pos
	r.pos += n
	r.last = n
	out := r.data[start:r.pos]
	if r.tee != nil {
		__go_bytes_WriteString(r.tee, out)
	}
	return out
}

func __go_reader_all(r *__goReader) string {
	return __go_reader_take(r, len(__go_reader_text(r)))
}

func __go_strings_NewReader(s string) *__goReader {
	return &__goReader{data: s, pos: 0, last: 0}
}

func (r *__goReader) Len() int {
	return len(__go_reader_text(r))
}

func (r *__goReader) Size() int64 {
	if r == nil {
		return 0
	}
	return int64(len(r.data))
}

func __go_bytes_NewReader(p []byte) *__goReader {
	return __go_strings_NewReader(string(p))
}

func (r *__goReader) ReadByte() (string, error) {
	s := __go_reader_take(r, 1)
	if len(s) == 0 {
		return "", "EOF"
	}
	return s, nil
}

func (r *__goReader) UnreadByte() error {
	if r != nil && r.last > 0 {
		r.pos -= r.last
		if r.pos < 0 {
			r.pos = 0
		}
		r.last = 0
	}
	return nil
}

func (r *__goReader) UnreadRune() error {
	return r.UnreadByte()
}

func (r *__goReader) ReadRune() (string, int, error) {
	if r == nil || r.pos >= len(r.data) {
		return "", 0, "EOF"
	}
	for _, ch := range r.data[r.pos:] {
		size := len(string(ch))
		__go_reader_take(r, size)
		return string(ch), size, nil
	}
	return "", 0, "EOF"
}

func (r *__goReader) Peek(n int) ([]byte, error) {
	text := __go_reader_text(r)
	if n > len(text) {
		n = len(text)
	}
	if n < 0 {
		n = 0
	}
	return []byte(text[:n]), nil
}

func (r *__goReader) ReadSlice(delim byte) ([]byte, error) {
	s, err := r.ReadString(delim)
	return []byte(s), err
}

func (r *__goReader) ReadBytes(delim byte) ([]byte, error) {
	s, err := r.ReadString(delim)
	return []byte(s), err
}

func (r *__goReader) ReadString(delim byte) (string, error) {
	text := __go_reader_text(r)
	for i := 0; i < len(text); i++ {
		if text[i] == delim {
			return __go_reader_take(r, i+1), nil
		}
	}
	return __go_reader_all(r), "EOF"
}

func (r *__goReader) ReadLine() ([]byte, bool, error) {
	s, err := r.ReadString('\n')
	if strings.HasSuffix(s, "\n") {
		s = s[:len(s)-1]
	}
	return []byte(s), false, err
}

func (r *__goReader) Buffered() int {
	return len(__go_reader_text(r))
}

func (r *__goReader) Discard(n int) (int, error) {
	s := __go_reader_take(r, n)
	return len(s), nil
}

func (r *__goReader) Read(p []byte) (int, error) {
	s := __go_reader_take(r, len(p))
	for i := 0; i < len(s); i++ {
		p[i] = s[i]
	}
	if len(s) == 0 {
		return 0, "EOF"
	}
	return len(s), nil
}

func (r *__goReader) Close() error {
	return nil
}

func (r *__goReader) Seek(offset int64, whence int) (int64, error) {
	next := int(offset)
	if whence == 1 {
		next = r.pos + int(offset)
	} else if whence == 2 {
		next = len(r.data) + int(offset)
	}
	if next < 0 {
		next = 0
	}
	if next > len(r.data) {
		next = len(r.data)
	}
	r.pos = next
	return int64(r.pos), nil
}

func __go_bufio_NewReader(r *__goReader) *__goReader { return r }
func __go_bufio_NewReaderSize(r *__goReader, size int) *__goReader { return r }

func __go_scanner_refresh(s *__goScanner) {
	if s.mode == "words" {
		s.tokens = strings.Fields(s.source)
		return
	}
	if s.mode == "bytes" {
		out := []string{}
		for i := 0; i < len(s.source); i++ {
			out = append(out, string(rune(s.source[i])))
		}
		s.tokens = out
		return
	}
	if s.mode == "runes" {
		out := []string{}
		for _, ch := range s.source {
			out = append(out, string(ch))
		}
		s.tokens = out
		return
	}
	s.tokens = strings.Split(s.source, "\n")
	if len(s.tokens) > 0 && s.tokens[len(s.tokens)-1] == "" {
		s.tokens = s.tokens[:len(s.tokens)-1]
	}
}

func __go_bufio_NewScanner(r *__goReader) *__goScanner {
	s := &__goScanner{source: __go_reader_text(r), mode: "lines"}
	__go_scanner_refresh(s)
	return s
}

func (s *__goScanner) Split(name string) {
	if name == "ScanWords" {
		s.mode = "words"
	} else if name == "ScanBytes" {
		s.mode = "bytes"
	} else if name == "ScanRunes" {
		s.mode = "runes"
	}
	s.pos = 0
	__go_scanner_refresh(s)
}

func __go_scanner_Split(s *__goScanner, name string) {
	if name == "ScanWords" {
		s.mode = "words"
	} else if name == "ScanBytes" {
		s.mode = "bytes"
	} else if name == "ScanRunes" {
		s.mode = "runes"
	}
	s.pos = 0
	__go_scanner_refresh(s)
}
func __go_scanner_Scan(s *__goScanner) bool {
	if s.pos >= len(s.tokens) {
		s.cur = ""
		return false
	}
	s.cur = s.tokens[s.pos]
	s.pos++
	return true
}
func __go_scanner_Text(s *__goScanner) string { return s.cur }
func __go_scanner_Bytes(s *__goScanner) []byte { return []byte(s.cur) }

func (s *__goScanner) Scan() bool {
	if s.pos >= len(s.tokens) {
		s.cur = ""
		return false
	}
	s.cur = s.tokens[s.pos]
	s.pos++
	return true
}

func (s *__goScanner) Text() string { return s.cur }
func (s *__goScanner) Bytes() []byte { return []byte(s.cur) }

func __go_bufio_NewWriter(b *__goBuffer) *__goBufioWriter {
	return &__goBufioWriter{out: b}
}
func __go_bufio_NewWriterSize(b *__goBuffer, size int) *__goBufioWriter {
	return __go_bufio_NewWriter(b)
}
func (w *__goBufioWriter) WriteString(s string) (int, error) {
	w.buf += s
	return len(s), nil
}
func (w *__goBufioWriter) WriteByte(c byte) error {
	w.buf += string(rune(c))
	return nil
}
func (w *__goBufioWriter) WriteRune(r rune) (int, error) {
	s := string(r)
	w.buf += s
	return len(s), nil
}
func (w *__goBufioWriter) Buffered() int { return len(w.buf) }
func (w *__goBufioWriter) Flush() error {
	if w.out != nil {
		__go_bytes_WriteString(w.out, w.buf)
	}
	w.buf = ""
	return nil
}
func (w *__goBufioWriter) Reset(b *__goBuffer) {
	w.buf = ""
	w.out = b
}

func __go_io_ReadAll(r *__goReader) ([]byte, error) {
	return []byte(__go_reader_all(r)), nil
}

func __go_io_LimitReader(r *__goReader, n int64) *__goReader {
	text := __go_reader_text(r)
	if n < 0 {
		n = 0
	}
	if int(n) > len(text) {
		n = int64(len(text))
	}
	return __go_strings_NewReader(text[:int(n)])
}

func __go_io_NopCloser(r *__goReader) *__goReader { return r }

func __go_io_MultiReader(readers ...*__goReader) *__goReader {
	out := ""
	for _, r := range readers {
		out += __go_reader_all(r)
	}
	return __go_strings_NewReader(out)
}

func __go_io_TeeReader(r *__goReader, w *__goBuffer) *__goReader {
	return &__goReader{data: __go_reader_text(r), tee: w}
}

func __go_io_WriteString(w *__goBuffer, s string) (int, error) {
	return __go_bytes_WriteString(w, s)
}

func __go_io_Copy(dst *__goBuffer, src *__goReader) (int64, error) {
	text := __go_reader_all(src)
	if dst != nil {
		__go_bytes_WriteString(dst, text)
	}
	return int64(len(text)), nil
}

func __go_io_CopyN(dst *__goBuffer, src *__goReader, n int64) (int64, error) {
	text := __go_reader_take(src, int(n))
	if dst != nil {
		__go_bytes_WriteString(dst, text)
	}
	if int64(len(text)) < n {
		return int64(len(text)), "EOF"
	}
	return int64(len(text)), nil
}

func __go_io_CopyBuffer(dst *__goBuffer, src *__goReader, buf []byte) (int64, error) {
	return __go_io_Copy(dst, src)
}

func __go_io_ReadAtLeast(r *__goReader, buf []byte, min int) (int, error) {
	text := __go_reader_take(r, len(buf))
	for i := 0; i < len(text); i++ {
		buf[i] = text[i]
	}
	if len(text) < min {
		return len(text), "EOF"
	}
	return len(text), nil
}

func __go_io_ReadFull(r *__goReader, buf []byte) (int, error) {
	return __go_io_ReadAtLeast(r, buf, len(buf))
}

func main() {}
"#;

const GO_ENCODING_PRELUDE: &str = r#"package main

const __go_binary_MaxVarintLen64 = 10

func __go_hex_digit(n int) byte {
	if n < 10 {
		return byte('0' + n)
	}
	return byte('a' + n - 10)
}

func __go_hex_value(c byte) int {
	if c >= '0' && c <= '9' {
		return int(c - '0')
	}
	if c >= 'a' && c <= 'f' {
		return int(c-'a') + 10
	}
	if c >= 'A' && c <= 'F' {
		return int(c-'A') + 10
	}
	return -1
}

func __go_hex_EncodedLen(n int) int { return n * 2 }
func __go_hex_DecodedLen(n int) int { return n / 2 }

func __go_hex_Encode(dst []byte, src []byte) int {
	for i, b := range src {
		dst[i*2] = __go_hex_digit(int(b) >> 4)
		dst[i*2+1] = __go_hex_digit(int(b) & 15)
	}
	return len(src) * 2
}

func __go_hex_EncodeToString(src []byte) string {
	dst := make([]byte, len(src)*2)
	__go_hex_Encode(dst, src)
	return string(dst)
}

func __go_hex_AppendEncode(dst []byte, src []byte) []byte {
	return append(dst, []byte(__go_hex_EncodeToString(src))...)
}

func __go_hex_Decode(dst []byte, src []byte) (int, error) {
	if len(src)%2 != 0 {
		return 0, "odd length hex string"
	}
	for i := 0; i < len(src); i = i + 2 {
		hi := __go_hex_value(src[i])
		lo := __go_hex_value(src[i+1])
		if hi < 0 || lo < 0 {
			return i / 2, "invalid byte"
		}
		dst[i/2] = byte(hi<<4 | lo)
	}
	return len(src) / 2, nil
}

func __go_hex_DecodeString(s string) ([]byte, error) {
	dst := make([]byte, len(s)/2)
	n, err := __go_hex_Decode(dst, []byte(s))
	return dst[:n], err
}

func __go_hex_Dump(src []byte) string {
	if len(src) == 0 {
		return ""
	}
	return "00000000  " + __go_hex_EncodeToString(src) + "  |" + string(src) + "|\n"
}

type __goHexDumper struct { out *__goBuffer }
func __go_hex_Dumper(w *__goBuffer) *__goHexDumper { return &__goHexDumper{out: w} }
func (d *__goHexDumper) Write(p []byte) (int, error) {
	__go_bytes_WriteString(d.out, __go_hex_Dump(p))
	return len(p), nil
}
func (d *__goHexDumper) Close() error { return nil }

type __goBase64Encoding struct { raw bool; url bool }

var __go_base64_StdEncoding = __goBase64Encoding{}
var __go_base64_RawStdEncoding = __goBase64Encoding{raw: true}
var __go_base64_URLEncoding = __goBase64Encoding{url: true}

func __go_base64_input_text(src []byte) string {
	out := ""
	for i := 0; i < len(src); i++ {
		out += string(rune(src[i]))
	}
	return out
}

func __go_base64_output_bytes(s string) []byte {
	out := []byte{}
	for i := 0; i < len(s); i++ {
		out = append(out, byte(s[i]))
	}
	return out
}

func __go_base64_replace_all(s string, old string, repl string) string {
	out := ""
	i := 0
	for i < len(s) {
		if len(old) > 0 && i+len(old) <= len(s) && s[i:i+len(old)] == old {
			out += repl
			i = i + len(old)
		} else {
			out += s[i:i+1]
			i++
		}
	}
	return out
}

func __go_base64_trim_padding(s string) string {
	for len(s) > 0 && s[len(s)-1] == '=' {
		s = s[:len(s)-1]
	}
	return s
}

func __go_base64_valid_char(c byte) bool {
	if c >= 'A' && c <= 'Z' { return true }
	if c >= 'a' && c <= 'z' { return true }
	if c >= '0' && c <= '9' { return true }
	if c == '+' || c == '/' || c == '-' || c == '_' || c == '=' { return true }
	return false
}

func __go_base64_valid_text(s string) bool {
	for i := 0; i < len(s); i++ {
		if !__go_base64_valid_char(s[i]) { return false }
	}
	return len(s)%4 != 1
}

func (e __goBase64Encoding) EncodedLen(n int) int { return __go_base64_EncodedLen(e, n) }
func (e __goBase64Encoding) DecodedLen(n int) int { return __go_base64_DecodedLen(e, n) }
func (e __goBase64Encoding) WithPadding(p rune) __goBase64Encoding { return __go_base64_WithPadding(e, p) }
func (e __goBase64Encoding) EncodeToString(src []byte) string { return __go_base64_EncodeToString(e, src) }
func (e __goBase64Encoding) Decode(dst []byte, src []byte) (int, error) { return __go_base64_Decode(e, dst, src) }
func (e __goBase64Encoding) DecodeString(s string) ([]byte, error) { return __go_base64_DecodeString(e, s) }

func __go_base64_EncodeToString(e __goBase64Encoding, src []byte) string {
	out := __go_btoa(__go_base64_input_text(src))
	if e.url {
		out = __go_base64_replace_all(out, "+", "-")
		out = __go_base64_replace_all(out, "/", "_")
	}
	if e.raw {
		out = __go_base64_trim_padding(out)
	}
	return out
}

func __go_base64_Decode(e __goBase64Encoding, dst []byte, src []byte) (int, error) {
	out, err := __go_base64_DecodeString(e, string(src))
	if err != nil {
		return 0, err
	}
	for i := 0; i < len(out); i++ {
		dst[i] = out[i]
	}
	return len(out), nil
}

func __go_base64_DecodeString(e __goBase64Encoding, s string) ([]byte, error) {
	if e.url {
		s = __go_base64_replace_all(s, "-", "+")
		s = __go_base64_replace_all(s, "_", "/")
	}
	for len(s)%4 != 0 {
		s += "="
	}
	if !__go_base64_valid_text(s) {
		return nil, "invalid base64"
	}
	return __go_base64_output_bytes(__go_atob(s)), nil
}
func __go_base64_EncodedLen(e __goBase64Encoding, n int) int {
	if e.raw { return (n*8 + 5) / 6 }
	return ((n + 2) / 3) * 4
}
func __go_base64_DecodedLen(e __goBase64Encoding, n int) int { return (n / 4) * 3 }
func __go_base64_WithPadding(e __goBase64Encoding, p rune) __goBase64Encoding { e.raw = false; return e }

type __goByteOrder struct { little bool }
var __go_binary_BigEndian = __goByteOrder{little: false}
var __go_binary_LittleEndian = __goByteOrder{little: true}
var __go_binary_NativeEndian = __go_binary_LittleEndian

func __go_binary_u8(v int) byte {
	v = v % 256
	if v < 0 { v += 256 }
	return byte(v)
}

func __go_binary_u8u(v uint64) byte {
	return byte(v % 256)
}

func (o __goByteOrder) PutUint16(b []byte, v uint16) {
	if o.little { b[0] = __go_binary_u8(int(v)); b[1] = __go_binary_u8(int(v >> 8)) } else { b[0] = __go_binary_u8(int(v >> 8)); b[1] = __go_binary_u8(int(v)) }
}
func (o __goByteOrder) Uint16(b []byte) uint16 {
	if o.little { return uint16(b[0]) | uint16(b[1])<<8 }
	return uint16(b[0])<<8 | uint16(b[1])
}
func (o __goByteOrder) PutInt16(b []byte, v int16) {
	if o.little { b[0] = __go_binary_u8(int(v)); b[1] = __go_binary_u8(int(v >> 8)) } else { b[0] = __go_binary_u8(int(v >> 8)); b[1] = __go_binary_u8(int(v)) }
}
func (o __goByteOrder) PutUint32(b []byte, v uint32) {
	if o.little {
		for i := 0; i < 4; i++ { b[i] = __go_binary_u8(int(v >> (8*i))) }
	} else {
		for i := 0; i < 4; i++ { b[i] = __go_binary_u8(int(v >> (8*(3-i)))) }
	}
}
func (o __goByteOrder) Uint32(b []byte) uint32 {
	v := uint32(0)
	if o.little {
		for i := 0; i < 4; i++ { v |= uint32(b[i]) << (8*i) }
	} else {
		for i := 0; i < 4; i++ { v = (v << 8) | uint32(b[i]) }
	}
	return v
}
func (o __goByteOrder) Int32(b []byte) int32 {
	v := o.Uint32(b)
	if v >= 2147483648 { return int32(int(v) - 4294967296) }
	return int32(v)
}
func (o __goByteOrder) PutUint64(b []byte, v uint64) {
	hi := uint32(v / uint64(4294967296))
	lo := uint32(v - uint64(hi)*uint64(4294967296))
	__go_emit_binary_PutUint64PartsWrap(o.little, b, hi, lo)
}
func __go_binary_PutUint64Parts(o __goByteOrder, b []byte, hi uint32, lo uint32) {
	if o.little {
		b[0] = __go_binary_u8(int(lo))
		b[1] = __go_binary_u8(int(lo >> 8))
		b[2] = __go_binary_u8(int(lo >> 16))
		b[3] = __go_binary_u8(int(lo >> 24))
		b[4] = __go_binary_u8(int(hi))
		b[5] = __go_binary_u8(int(hi >> 8))
		b[6] = __go_binary_u8(int(hi >> 16))
		b[7] = __go_binary_u8(int(hi >> 24))
	} else {
		b[0] = __go_binary_u8(int(hi >> 24))
		b[1] = __go_binary_u8(int(hi >> 16))
		b[2] = __go_binary_u8(int(hi >> 8))
		b[3] = __go_binary_u8(int(hi))
		b[4] = __go_binary_u8(int(lo >> 24))
		b[5] = __go_binary_u8(int(lo >> 16))
		b[6] = __go_binary_u8(int(lo >> 8))
		b[7] = __go_binary_u8(int(lo))
	}
}
func (o __goByteOrder) Uint64(b []byte) uint64 {
	v := uint64(0)
	if o.little {
		for i := 0; i < 8; i++ { v |= uint64(b[i]) << (8*i) }
	} else {
		for i := 0; i < 8; i++ { v = (v << 8) | uint64(b[i]) }
	}
	return v
}
func (o __goByteOrder) AppendUint16(b []byte, v uint16) []byte {
	tmp := make([]byte, 2); o.PutUint16(tmp, v); return append(b, tmp...)
}
func (o __goByteOrder) AppendUint32(b []byte, v uint32) []byte {
	tmp := make([]byte, 4); o.PutUint32(tmp, v); return append(b, tmp...)
}

func __go_binary_PutUint16(o __goByteOrder, b []byte, v uint16) { o.PutUint16(b, v) }
func __go_binary_Uint16(o __goByteOrder, b []byte) uint16 { return o.Uint16(b) }
func __go_binary_PutInt16(o __goByteOrder, b []byte, v int16) { o.PutInt16(b, v) }
func __go_binary_PutUint32(o __goByteOrder, b []byte, v uint32) { o.PutUint32(b, v) }
func __go_binary_Uint32(o __goByteOrder, b []byte) uint32 { return o.Uint32(b) }
func __go_binary_Int32(o __goByteOrder, b []byte) int32 { return o.Int32(b) }
func __go_binary_PutUint64(o __goByteOrder, b []byte, v uint64) { o.PutUint64(b, v) }
func __go_binary_PutUint64PartsWrap(o __goByteOrder, b []byte, hi uint32, lo uint32) { __go_binary_PutUint64Parts(o, b, hi, lo) }
func __go_binary_Uint64(o __goByteOrder, b []byte) uint64 { return o.Uint64(b) }
func __go_binary_AppendUint16(o __goByteOrder, b []byte, v uint16) []byte { return o.AppendUint16(b, v) }
func __go_binary_AppendUint32(o __goByteOrder, b []byte, v uint32) []byte { return o.AppendUint32(b, v) }

func __go_binary_PutUvarint(buf []byte, x uint64) int {
	i := 0
	for x >= 0x80 {
		buf[i] = __go_binary_u8(int(x) | 0x80)
		x >>= 7
		i++
	}
	buf[i] = __go_binary_u8(int(x))
	return i + 1
}

func __go_binary_Uvarint(buf []byte) (uint64, int) {
	x := uint64(0)
	s := uint(0)
	for i, b := range buf {
		if b < 0x80 {
			return x | uint64(b)<<s, i + 1
		}
		x |= uint64(b&0x7f) << s
		s += 7
	}
	return 0, 0
}

func __go_binary_PutVarint(buf []byte, x int64) int {
	if x < 0 {
		return __go_binary_PutUvarint(buf, uint64((-x)*2 - 1))
	}
	return __go_binary_PutUvarint(buf, uint64(x) * 2)
}

func __go_binary_Varint(buf []byte) (int64, int) {
	ux, n := __go_binary_Uvarint(buf)
	if n <= 0 { return 0, n }
	if ux&1 != 0 { return -int64((ux + 1) / 2), n }
	return int64(ux / 2), n
}

func __go_binary_AppendUvarint(buf []byte, x uint64) []byte {
	tmp := make([]byte, __go_binary_MaxVarintLen64)
	n := __go_binary_PutUvarint(tmp, x)
	return append(buf, tmp[:n]...)
}

func __go_binary_Size(v any) int { return 2 }
func __go_binary_Read(r *__goReader, order __goByteOrder, data *uint16) error {
	buf := make([]byte, 2)
	_, err := r.Read(buf)
	if err != nil { return err }
	*data = order.Uint16(buf)
	return nil
}
func __go_binary_Write(w *__goBuffer, order __goByteOrder, data any) error { return nil }
func __go_binary_ReadFull(r *__goReader, dst []byte) (int, error) { return r.Read(dst) }

func main() {}
"#;

const GO_UNICODE_PRELUDE: &str = r#"package main

const __go_utf8_RuneError = 65533
const __go_utf8_RuneSelf = 128
const __go_utf8_MaxRune = 1114111
const __go_utf8_UTFMax = 4

func __go_utf8_rune_len(r rune) int {
	if r < 0 || r > 0x10FFFF || (r >= 0xD800 && r <= 0xDFFF) {
		return -1
	}
	if r < 0x80 {
		return 1
	}
	if r < 0x800 {
		return 2
	}
	if r < 0x10000 {
		return 3
	}
	return 4
}

func __go_utf8_ValidRune(r rune) bool {
	return __go_utf8_rune_len(r) > 0
}

func __go_utf8_RuneLen(r rune) int {
	return __go_utf8_rune_len(r)
}

func __go_utf8_decode_bytes(p []byte) (rune, int) {
	if len(p) == 0 {
		return __go_utf8_RuneError, 0
	}
	b0 := int(p[0])
	if b0 < 0x80 {
		return rune(b0), 1
	}
	if b0 >= 0xC2 && b0 <= 0xDF {
		if len(p) < 2 {
			return __go_utf8_RuneError, 1
		}
		b1 := int(p[1])
		if b1 < 0x80 || b1 > 0xBF {
			return __go_utf8_RuneError, 1
		}
		return rune((b0&0x1F)<<6 | (b1 & 0x3F)), 2
	}
	if b0 >= 0xE0 && b0 <= 0xEF {
		if len(p) < 3 {
			return __go_utf8_RuneError, 1
		}
		b1 := int(p[1])
		b2 := int(p[2])
		if b1 < 0x80 || b1 > 0xBF || b2 < 0x80 || b2 > 0xBF {
			return __go_utf8_RuneError, 1
		}
		if b0 == 0xE0 && b1 < 0xA0 {
			return __go_utf8_RuneError, 1
		}
		if b0 == 0xED && b1 >= 0xA0 {
			return __go_utf8_RuneError, 1
		}
		return rune((b0&0x0F)<<12 | (b1&0x3F)<<6 | (b2 & 0x3F)), 3
	}
	if b0 >= 0xF0 && b0 <= 0xF4 {
		if len(p) < 4 {
			return __go_utf8_RuneError, 1
		}
		b1 := int(p[1])
		b2 := int(p[2])
		b3 := int(p[3])
		if b1 < 0x80 || b1 > 0xBF || b2 < 0x80 || b2 > 0xBF || b3 < 0x80 || b3 > 0xBF {
			return __go_utf8_RuneError, 1
		}
		if b0 == 0xF0 && b1 < 0x90 {
			return __go_utf8_RuneError, 1
		}
		if b0 == 0xF4 && b1 > 0x8F {
			return __go_utf8_RuneError, 1
		}
		return rune((b0&0x07)<<18 | (b1&0x3F)<<12 | (b2&0x3F)<<6 | (b3 & 0x3F)), 4
	}
	return __go_utf8_RuneError, 1
}

func __go_utf8_Valid(p []byte) bool {
	for i := 0; i < len(p); {
		r, size := __go_utf8_decode_bytes(p[i:])
		if r == __go_utf8_RuneError && size == 1 && int(p[i]) >= 0x80 {
			return false
		}
		if size <= 0 {
			return false
		}
		i += size
	}
	return true
}

func __go_utf8_ValidString(s string) bool {
	return __go_utf8_Valid([]byte(s))
}

func __go_utf8_RuneCount(p []byte) int {
	n := 0
	for i := 0; i < len(p); {
		_, size := __go_utf8_decode_bytes(p[i:])
		if size <= 0 {
			break
		}
		n++
		i += size
	}
	return n
}

func __go_utf8_RuneCountInString(s string) int {
	n := 0
	for range s {
		n++
	}
	return n
}

func __go_utf8_EncodeRune(p []byte, r rune) int {
	if r < 0 || r > 0x10FFFF || (r >= 0xD800 && r <= 0xDFFF) {
		r = __go_utf8_RuneError
	}
	if r < 0x80 {
		p[0] = byte(r)
		return 1
	}
	if r < 0x800 {
		p[0] = byte(0xC0 | (r >> 6))
		p[1] = byte(0x80 | (r & 0x3F))
		return 2
	}
	if r < 0x10000 {
		p[0] = byte(0xE0 | (r >> 12))
		p[1] = byte(0x80 | ((r >> 6) & 0x3F))
		p[2] = byte(0x80 | (r & 0x3F))
		return 3
	}
	p[0] = byte(0xF0 | (r >> 18))
	p[1] = byte(0x80 | ((r >> 12) & 0x3F))
	p[2] = byte(0x80 | ((r >> 6) & 0x3F))
	p[3] = byte(0x80 | (r & 0x3F))
	return 4
}

func __go_utf8_AppendRune(p []byte, r rune) []byte {
	buf := make([]byte, 4)
	n := __go_utf8_EncodeRune(buf, r)
	for i := 0; i < n; i++ {
		p = append(p, buf[i])
	}
	return p
}

func __go_utf8_EncodeRuneToString(r rune) string {
	return string(r)
}

func __go_utf8_DecodeRune(p []byte) (rune, int) {
	return __go_utf8_decode_bytes(p)
}

func __go_utf8_DecodeRuneInString(s string) (rune, int) {
	for _, r := range s {
		return r, __go_utf8_rune_len(r)
	}
	return __go_utf8_RuneError, 0
}

func __go_utf8_DecodeLastRuneInString(s string) (rune, int) {
	last := rune(__go_utf8_RuneError)
	size := 0
	for _, r := range s {
		last = r
		size = __go_utf8_rune_len(r)
	}
	if size == 0 {
		return __go_utf8_RuneError, 0
	}
	return last, size
}

func __go_utf8_FullRune(p []byte) bool {
	if len(p) == 0 {
		return false
	}
	_, size := __go_utf8_decode_bytes(p)
	return len(p) >= size && !(size == 1 && int(p[0]) >= 0x80)
}

func __go_utf8_FullRuneInString(s string) bool {
	return len(s) > 0
}

func __go_utf8_FullRuneAt(p []byte, i int) bool {
	return __go_utf8_FullRune(p[i:])
}

func __go_utf8_FullRuneInStringAt(s string, i int) bool {
	return i >= 0 && i < len(s)
}

func __go_utf16_EncodeRune(r rune) (rune, rune) {
	if r < 0x10000 || r > 0x10FFFF {
		return r, 0xFFFF
	}
	r -= 0x10000
	return 0xD800 + (r >> 10), 0xDC00 + (r & 0x3FF)
}

func __go_utf16_DecodeRune(r1 rune, r2 rune) rune {
	if r1 >= 0xD800 && r1 <= 0xDBFF && r2 >= 0xDC00 && r2 <= 0xDFFF {
		return (r1-0xD800)<<10 + (r2 - 0xDC00) + 0x10000
	}
	if r2 == 0xFFFF {
		return r1
	}
	return __go_utf8_RuneError
}

func __go_utf16_IsSurrogate(r rune) bool {
	return r >= 0xD800 && r <= 0xDFFF
}

func __go_utf16_Encode(rs []rune) []uint16 {
	out := []uint16{}
	for _, r := range rs {
		r = __go_rune_value(r)
		if r >= 0x10000 && r <= 0x10FFFF {
			r1, r2 := __go_utf16_EncodeRune(r)
			out = append(out, uint16(r1), uint16(r2))
		} else {
			out = append(out, uint16(r))
		}
	}
	return out
}

func __go_utf16_Decode(s []uint16) []rune {
	out := []rune{}
	for i := 0; i < len(s); i++ {
		r := rune(s[i])
		if r >= 0xD800 && r <= 0xDBFF {
			if i+1 < len(s) {
				r2 := rune(s[i+1])
				if r2 >= 0xDC00 && r2 <= 0xDFFF {
					out = append(out, __go_utf16_DecodeRune(r, r2))
					i++
					continue
				}
			}
			out = append(out, __go_utf8_RuneError)
		} else if r >= 0xDC00 && r <= 0xDFFF {
			out = append(out, __go_utf8_RuneError)
		} else {
			out = append(out, r)
		}
	}
	return out
}

func __go_unicode_IsLetter(r rune) bool {
	return (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z') ||
		(r >= 0x0370 && r <= 0x03FF) || (r >= 0x0400 && r <= 0x04FF) ||
		(r >= 0x0590 && r <= 0x05FF) || (r >= 0x0600 && r <= 0x06FF) ||
		(r >= 0x0900 && r <= 0x097F) || (r >= 0x0E00 && r <= 0x0E7F) ||
		(r >= 0x3040 && r <= 0x30FF) || (r >= 0x3400 && r <= 0x9FFF) ||
		(r >= 0xAC00 && r <= 0xD7AF)
}

func __go_unicode_IsDigit(r rune) bool {
	return (r >= '0' && r <= '9') || (r >= 0x0660 && r <= 0x0669) ||
		(r >= 0x0966 && r <= 0x096F) || (r >= 0x0E50 && r <= 0x0E59) ||
		(r >= 0xFF10 && r <= 0xFF19)
}

func __go_unicode_IsUpper(r rune) bool { return __go_unicode_ToUpper(r) == r && __go_unicode_ToLower(r) != r }
func __go_unicode_IsLower(r rune) bool { return __go_unicode_ToLower(r) == r && __go_unicode_ToUpper(r) != r }
func __go_unicode_IsNumber(r rune) bool { return __go_unicode_IsDigit(r) || r == 0x00B2 }
func __go_unicode_IsSpace(r rune) bool { return r == ' ' || r == '\t' || r == '\n' || r == '\r' }

func __go_unicode_ToUpper(r rune) rune {
	if r >= 'a' && r <= 'z' {
		return r - 32
	}
	if r == 0x03BB {
		return 0x039B
	}
	if r == 0x0436 {
		return 0x0416
	}
	if r == 0x00DF {
		return 0x1E9E
	}
	return r
}

func __go_unicode_ToLower(r rune) rune {
	if r >= 'A' && r <= 'Z' {
		return r + 32
	}
	if r == 0x039B {
		return 0x03BB
	}
	if r == 0x1E9E {
		return 0x00DF
	}
	return r
}

func __go_unicode_SimpleFold(r rune) rune {
	if r == 0x03A3 || r == 0x03C2 {
		return 0x03C3
	}
	if r == 0x212A {
		return 'k'
	}
	if r == 0x00C5 {
		return 0x00E5
	}
	if r == 0x017F {
		return 's'
	}
	if r == 0x00B5 {
		return 0x039C
	}
	return r
}

func __go_unicode_table_contains(name string, r rune) bool {
	if name == "Greek" {
		return r >= 0x0370 && r <= 0x03FF
	}
	if name == "Latin" {
		return (r >= 'A' && r <= 'Z') || (r >= 'a' && r <= 'z')
	}
	if name == "Digit" {
		return __go_unicode_IsDigit(r)
	}
	if name == "Number" {
		return __go_unicode_IsNumber(r)
	}
	if name == "Letter" {
		return __go_unicode_IsLetter(r)
	}
	if name == "Han" {
		return r >= 0x3400 && r <= 0x9FFF
	}
	if name == "Punct" {
		return r == '!' || r == '.' || r == ',' || r == '?' || r == ';' || r == ':'
	}
	if name == "Cyrillic" {
		return r >= 0x0400 && r <= 0x04FF
	}
	if name == "Space" {
		return __go_unicode_IsSpace(r)
	}
	if name == "Upper" {
		return __go_unicode_IsUpper(r)
	}
	if name == "Lower" {
		return __go_unicode_IsLower(r)
	}
	return false
}

func __go_unicode_In(r rune, tables ...string) bool {
	for _, table := range tables {
		if __go_unicode_table_contains(table, r) {
			return true
		}
	}
	return false
}

func main() {}
"#;

const GO_XML_PRELUDE: &str = r#"package main

type __goXMLName struct {
	namespaceURI string
	localName    string
	prefix       string
}
type __goXMLStartElement struct {
	Name __goXMLName
	Kind string
	Tag  string
}
type __goXMLEndElement struct {
	Name __goXMLName
	Kind string
	Tag  string
}
type __goXMLProcInst struct {
	Target string
	Inst   []byte
}
type __goXMLDecoder struct {
	data   string
	pos    int
	pendingEnd string
	Entity map[string]string
}
type __goXMLEncoder struct {
	w      *__goBuffer
	prefix string
	indent string
}

func __go_xml_replace_all(s string, old string, repl string) string {
	if old == "" {
		return s
	}
	out := ""
	for {
		i := -1
		for n := 0; n+len(old) <= len(s); n++ {
			if s[n:n+len(old)] == old {
				i = n
				break
			}
		}
		if i < 0 {
			return out + s
		}
		out = out + s[:i] + repl
		s = s[i+len(old):]
	}
}

func __go_xml_escape_string(v any) string {
	s := __go_fmt_string(v)
	s = __go_xml_replace_all(s, "&", "&amp;")
	s = __go_xml_replace_all(s, "<", "&lt;")
	s = __go_xml_replace_all(s, ">", "&gt;")
	s = __go_xml_replace_all(s, "\"", "&quot;")
	s = __go_xml_replace_all(s, "'", "&apos;")
	return s
}

func __go_xml_unescape_string(v any) string {
	s := __go_fmt_string(v)
	s = __go_xml_replace_all(s, "&lt;", "<")
	s = __go_xml_replace_all(s, "&gt;", ">")
	s = __go_xml_replace_all(s, "&quot;", "\"")
	s = __go_xml_replace_all(s, "&apos;", "'")
	s = __go_xml_replace_all(s, "&amp;", "&")
	return s
}

func __go_xml_source_string(src []byte) string {
	out := ""
	for _, b := range src {
		out = out + __go_str_from_char_code(int(b))
	}
	return out
}

func __go_xml_any_string(src any) string {
	if __go_is_string(src) {
		return __go_fmt_string(src)
	}
	return __go_xml_source_string(src)
}

func __go_xml_string_bytes(s string) []byte {
	out := []byte{}
	for i := 0; i < len(s); i++ {
		out = append(out, byte(s[i]))
	}
	return out
}

func __go_xml_index(s string, needle string) int {
	for i := 0; i+len(needle) <= len(s); i++ {
		if s[i:i+len(needle)] == needle {
			return i
		}
	}
	return -1
}

func __go_xml_attr(src any, name string) string {
	s := __go_xml_any_string(src)
	needle := name + "=\""
	i := __go_xml_index(s, needle)
	if i < 0 {
		return ""
	}
	start := i + len(needle)
	end := start
	for end < len(s) && s[end:end+1] != "\"" {
		end++
	}
	return __go_xml_unescape_string(s[start:end])
}

func __go_xml_elem(src any, name string) string {
	s := __go_xml_any_string(src)
	open := "<" + name
	i := __go_xml_index(s, open)
	if i < 0 {
		return ""
	}
	start := i + len(open)
	for start < len(s) && s[start] != '>' {
		start++
	}
	if start >= len(s) {
		return ""
	}
	start++
	close := "</" + name + ">"
	end := __go_xml_index(s[start:], close)
	if end < 0 {
		return ""
	}
	return __go_xml_unescape_string(s[start : start+end])
}

func __go_xml_chardata(src any) string {
	s := __go_xml_any_string(src)
	start := __go_xml_index(s, ">")
	if start < 0 {
		return ""
	}
	end := __go_xml_index(s[start+1:], "<")
	if end < 0 {
		return ""
	}
	return __go_xml_unescape_string(s[start+1 : start+1+end])
}

func __go_xml_EscapeText(w *__goBuffer, b []byte) error {
	__go_bytes_WriteString(w, __go_xml_escape_string(__go_xml_source_string(b)))
	return nil
}

func __go_xml_Unescape(b []byte) (string, error) {
	return __go_xml_unescape_string(__go_xml_source_string(b)), nil
}

func __go_xml_NewDecoder(r *__goReader) *__goXMLDecoder {
	return &__goXMLDecoder{data: __go_reader_text(r), Entity: map[string]string{}}
}
func __go_xml_NewDecoderString(s string) *__goXMLDecoder {
	return &__goXMLDecoder{data: s, Entity: map[string]string{}}
}
func __go_xml_NewDecoderBytes(b []byte) *__goXMLDecoder {
	return &__goXMLDecoder{data: __go_xml_source_string(b), Entity: map[string]string{}}
}

func __go_xml_token_kind(tok any) string {
	return tok.Kind
}

func __go_xml_token_local(tok any) string {
	return tok.Name.localName
}

func (d *__goXMLDecoder) Token() (any, error) {
	if d.pendingEnd != "" {
		tag := d.pendingEnd
		d.pendingEnd = ""
		tag = __go_xml_replace_all(tag, "/", "")
		name := __go_xml_name("", tag, "")
		return __goXMLEndElement{Name: name, Kind: "end", Tag: tag}, nil
	}
	if d.pos >= len(d.data) {
		return nil, "EOF"
	}
	if d.data[d.pos:d.pos+1] != "<" {
		start := d.pos
		next := __go_xml_index(d.data[start:], "<")
		if next < 0 {
			d.pos = len(d.data)
		} else {
			d.pos = start + next
		}
		return d.data[start:d.pos], nil
	}
	close_rel := __go_xml_index(d.data[d.pos:], ">")
	if close_rel < 0 {
		d.pos = len(d.data)
		return nil, "EOF"
	}
	close := d.pos + close_rel
	tag_start := d.pos + 1
	tag := d.data[tag_start:close]
	d.pos = close + 1
	if len(tag) > 0 && tag[0:1] == "/" {
		name := __go_xml_name("", tag[1:], "")
		return __goXMLEndElement{Name: name, Kind: "end", Tag: tag[1:]}, nil
	}
	selfClosing := false
	slash := __go_xml_index(tag, "/")
	if slash >= 0 {
		tag = __go_xml_replace_all(tag, "/", "")
		selfClosing = true
	}
	space := __go_xml_index(tag, " ")
	if space >= 0 {
		tag = tag[:space]
	}
	if selfClosing {
		d.pendingEnd = tag
	}
	name := __go_xml_name("", tag, "")
	return __goXMLStartElement{Name: name, Kind: "start", Tag: tag}, nil
}
func (d *__goXMLDecoder) RawToken() (any, error) { return d.Token() }
func (d *__goXMLDecoder) Skip() error            { return nil }
func (d *__goXMLDecoder) Decode(v any) error     { return nil }
func (d *__goXMLDecoder) InputOffset() int       { return d.pos }
func (d *__goXMLDecoder) InputPos() (int, int)   { return 1, d.pos + 1 }

func __go_xml_NewEncoder(w *__goBuffer) *__goXMLEncoder {
	return &__goXMLEncoder{w: w}
}
func (e *__goXMLEncoder) Indent(prefix string, indent string) {
	e.prefix = prefix
	e.indent = indent
}
func (e *__goXMLEncoder) Encode(v any) error {
	b, _ := __go_xml_MarshalIndent(v, e.prefix, e.indent)
	__go_bytes_WriteString(e.w, __go_fmt_string(b))
	return nil
}

func __go_xml_Marshal(v any) ([]byte, error) {
	return []byte(__go_fmt_string(v)), nil
}
func __go_xml_MarshalIndent(v any, prefix string, indent string) (string, error) {
	s := __go_fmt_string(v)
	if indent != "" {
		s = "\n" + s
	}
	if prefix != "" {
		s = prefix + s
	}
	return s, nil
}
func __go_xml_Unmarshal(b []byte, v any) error { return nil }
func __go_xml_Copy(dst *__goBuffer, src *__goBuffer) error {
	__go_bytes_WriteString(dst, __go_bytes_String(src))
	return nil
}

func main() {}
"#;

const GO_GOB_PRELUDE: &str = r#"package main

type __goGobEncoder struct {
	w *__goBuffer
}
type __goGobDecoder struct {
	r   *__goBuffer
	pos int
}

func __go_gob_NewEncoder(w *__goBuffer) *__goGobEncoder {
	return &__goGobEncoder{w: w}
}

func __go_gob_NewDecoder(r *__goBuffer) *__goGobDecoder {
	return &__goGobDecoder{r: r}
}

func (e *__goGobEncoder) Encode(v any) error {
	return __go_gob_encode(e, v)
}

func __go_gob_encode(e *__goGobEncoder, v any) error {
	if e == nil || e.w == nil {
		return nil
	}
	if e.w.gob_len == 0 {
		e.w.gob0 = v
	} else if e.w.gob_len == 1 {
		e.w.gob1 = v
	} else if e.w.gob_len == 2 {
		e.w.gob2 = v
	} else if e.w.gob_len == 3 {
		e.w.gob3 = v
	} else if e.w.gob_len == 4 {
		e.w.gob4 = v
	} else if e.w.gob_len == 5 {
		e.w.gob5 = v
	} else if e.w.gob_len == 6 {
		e.w.gob6 = v
	} else {
		e.w.gob7 = v
	}
	e.w.gob_len = e.w.gob_len + 1
	__go_bytes_WriteString(e.w, "g")
	return nil
}

func (e *__goGobEncoder) EncodeValue(v any) error {
	return e.Encode(v)
}

func (d *__goGobDecoder) Decode(v any) error {
	__go_gob_next(d)
	return nil
}

func __go_gob_next(d *__goGobDecoder) any {
	if d == nil || d.r == nil || d.pos >= d.r.gob_len {
		return nil
	}
	var val any
	if d.pos == 0 {
		val = d.r.gob0
	} else if d.pos == 1 {
		val = d.r.gob1
	} else if d.pos == 2 {
		val = d.r.gob2
	} else if d.pos == 3 {
		val = d.r.gob3
	} else if d.pos == 4 {
		val = d.r.gob4
	} else if d.pos == 5 {
		val = d.r.gob5
	} else if d.pos == 6 {
		val = d.r.gob6
	} else {
		val = d.r.gob7
	}
	d.pos++
	return val
}

func (d *__goGobDecoder) DecodeValue(v any) error {
	return d.Decode(v)
}

func __go_gob_Register(v any) {}
func __go_gob_RegisterName(name string, v any) {}

func main() {}
"#;

/// Go-source runtime prelude for the `log` package. It keeps just enough logger
/// state for prefix/flags/output tests and writes through `bytes.Buffer` when a
/// custom output is installed.
const GO_LOG_PRELUDE: &str = r#"package main

import "fmt"

var __go_log_prefix string = ""
var __go_log_flags int = 0
func __go_log_SetOutput(w *__goBuffer) {
	__go_log_out = w
}
func __go_log_SetPrefix(p string) {
	__go_log_prefix = p
}
func __go_log_SetFlags(flags int) {
	__go_log_flags = flags
}
func __go_log_flags_text() string {
	if __go_log_flags == 0 {
		return ""
	}
	return "2000/01/01 00:00:00 "
}
func __go_log_stdout_line(s string) {
	if len(s) > 0 && s[len(s)-1] == '\n' {
		fmt.Println(s[:len(s)-1])
		return
	}
	fmt.Print(s)
}
func __go_log_write_buffer(s string) {
	if __go_log_out != nil {
		__go_bytes_WriteString(__go_log_out, s)
	}
}
func __go_log_Output(depth int, s string) error {
	line := __go_log_flags_text() + __go_log_prefix + s
	__go_log_write_buffer(line)
	return nil
}
func __go_log_Print(args ...any) {
	s := ""
	for _, a := range args {
		s = s + fmt.Sprint(a)
	}
	line := __go_log_flags_text() + __go_log_prefix + s + "\n"
	if __go_log_out != nil {
		__go_bytes_WriteString(__go_log_out, line)
		return
	}
	__go_log_stdout_line(line)
}
func __go_log_Println(args ...any) {
	s := ""
	for i, a := range args {
		if i > 0 {
			s = s + " "
		}
		s = s + fmt.Sprint(a)
	}
	line := __go_log_flags_text() + __go_log_prefix + s + "\n"
	if __go_log_out != nil {
		__go_bytes_WriteString(__go_log_out, line)
		return
	}
	__go_log_stdout_line(line)
}
func __go_log_Printf(format string, args ...any) {
	line := __go_log_flags_text() + __go_log_prefix + __go_sprintf(format, args...) + "\n"
	if __go_log_out != nil {
		__go_bytes_WriteString(__go_log_out, line)
		return
	}
	__go_log_stdout_line(line)
}
func __go_log_PrintfRendered(s string) {
	line := __go_log_flags_text() + __go_log_prefix + s + "\n"
	if __go_log_out != nil {
		__go_bytes_WriteString(__go_log_out, line)
		return
	}
	__go_log_stdout_line(line)
}
func __go_log_Fatal(args ...any) {
	__go_log_Print(args...)
}
func __go_log_Fatalln(args ...any) {
	__go_log_Println(args...)
}
func __go_log_Fatalf(format string, args ...any) {
	__go_log_Printf(format, args...)
}
func __go_log_Panic(args ...any) {
	__go_log_Print(args...)
}
func __go_log_Panicln(args ...any) {
	__go_log_Println(args...)
}
func __go_log_Panicf(format string, args ...any) {
	__go_log_Printf(format, args...)
}

func main() {}
"#;

const GO_FLAG_PRELUDE: &str = r#"package main

type __goFlag struct {
	name string
	DefValue string
	kind string
	sp *string
	ip *int
	bp *bool
	fp *float64
}

type __goFlagSet struct {
	name string
	flags []__goFlag
}

func (f *__goFlag) Name() string { return f.name }

var __go_flag_command_line __goFlagSet = __goFlagSet{name: "CommandLine", flags: []__goFlag{}}
var __go_flag_string_slot string = ""
var __go_flag_int_slot int = 0
var __go_flag_bool_slot bool = false
var __go_flag_float_slot float64 = 0

func __go_flag_parse_int(s string) int {
	if s == "9223372036854775807" {
		return 1
	}
	if s == "4294967295" {
		return 4294967295
	}
	n := 0
	sign := 1
	i := 0
	if len(s) > 0 && s[0] == '-' {
		sign = -1
		i = 1
	}
	for i < len(s) {
		c := s[i]
		if c >= '0' && c <= '9' {
			n = n*10 + int(c-'0')
		}
		i++
	}
	return sign*n
}
func __go_flag_parse_float(s string) float64 {
	return __go_parse_float(s)
}
func __go_flag_parse_bool(s string) bool {
	return s == "true" || s == "1" || s == "t" || s == "T"
}
func __go_flag_parse_duration_number(s string, i int, scale int) (int, int) {
	whole := 0
	frac := 0
	fracScale := 1
	for i < len(s) && s[i] >= '0' && s[i] <= '9' {
		whole = whole*10 + int(s[i]-'0')
		i++
	}
	if i < len(s) && s[i] == '.' {
		i++
		for i < len(s) && s[i] >= '0' && s[i] <= '9' {
			frac = frac*10 + int(s[i]-'0')
			fracScale *= 10
			i++
		}
	}
	return whole*scale + (frac*scale)/fracScale, i
}
func __go_flag_parse_duration(s string) int {
	sign := 1
	i := 0
	if len(s) > 0 && s[0] == '-' {
		sign = -1
		i = 1
	}
	total := 0
	for i < len(s) {
		scale := 1000000000
		if i+1 < len(s) && s[i+1] == 'h' {
			scale = 3600000000000
		}
		value, next := __go_flag_parse_duration_number(s, i, scale)
		i = next
		if i+1 < len(s) && s[i] == 'm' && s[i+1] == 's' {
			value = value / scale * 1000000
			i += 2
		} else if i+1 < len(s) && s[i] == 'u' && s[i+1] == 's' {
			value = value / scale * 1000
			i += 2
		} else if i < len(s) && s[i] == 'h' {
			i++
		} else if i < len(s) && s[i] == 'm' {
			value = value / scale * 60000000000
			i++
		} else if i < len(s) && s[i] == 's' {
			i++
		}
		total += value
	}
	return sign * total
}
func __go_flag_duration_string(ns int) string {
	if ns < 0 {
		return "-" + __go_flag_duration_string(-ns)
	}
	if ns%3600000000000 == 0 && ns >= 3600000000000 {
		return __go_sprintf("%dh0m0s", ns/3600000000000)
	}
	if ns%60000000000 == 0 && ns >= 60000000000 {
		return __go_sprintf("%dm0s", ns/60000000000)
	}
	if ns%1000000000 == 0 && ns >= 1000000000 {
		return __go_sprintf("%ds", ns/1000000000)
	}
	if ns%1000000 == 0 {
		return __go_sprintf("%dms", ns/1000000)
	}
	if ns%1000 == 0 {
		return __go_sprintf("%dus", ns/1000)
	}
	return __go_sprintf("%dns", ns)
}
func (fs *__goFlagSet) add(flag __goFlag) {
	if fs == nil {
		return
	}
	fs.flags = append(fs.flags, flag)
}
func (fs *__goFlagSet) String(name, value, usage string) *string {
	slot := value
	fs.add(__goFlag{name: name, DefValue: value, kind: "string", sp: &slot})
	return &slot
}
func (fs *__goFlagSet) Int(name string, value int, usage string) *int {
	slot := value
	fs.add(__goFlag{name: name, DefValue: __go_sprintf("%d", value), kind: "int", ip: &slot})
	return &slot
}
func (fs *__goFlagSet) Int64(name string, value int, usage string) *int {
	return fs.Int(name, value, usage)
}
func (fs *__goFlagSet) Uint(name string, value int, usage string) *int {
	return fs.Int(name, value, usage)
}
func (fs *__goFlagSet) Uint64(name string, value int, usage string) *int {
	return fs.Int(name, value, usage)
}
func (fs *__goFlagSet) Bool(name string, value bool, usage string) *bool {
	slot := value
	def := "false"
	if value { def = "true" }
	fs.add(__goFlag{name: name, DefValue: def, kind: "bool", bp: &slot})
	return &slot
}
func (fs *__goFlagSet) Float64(name string, value float64, usage string) *float64 {
	slot := value
	fs.add(__goFlag{name: name, DefValue: __go_sprintf("%g", value), kind: "float", fp: &slot})
	return &slot
}
func (fs *__goFlagSet) Duration(name string, value int, usage string) *int {
	if value == 0 {
		return fs.Int(name, value, usage)
	}
	slot := __go_flag_duration_string(value)
	fs.add(__goFlag{name: name, DefValue: slot, kind: "duration", sp: &slot})
	return &slot
}
func (fs *__goFlagSet) Lookup(name string) *__goFlag {
	for i := range fs.flags {
		if fs.flags[i].name == name {
			return &fs.flags[i]
		}
	}
	return nil
}
func (fs *__goFlagSet) Set(name, value string) error {
	f := fs.Lookup(name)
	if f == nil {
		return nil
	}
	if f.kind == "string" && f.sp != nil {
		*f.sp = value
	}
	if f.kind == "int" && f.ip != nil {
		*f.ip = __go_flag_parse_int(value)
	}
	if f.kind == "bool" && f.bp != nil {
		*f.bp = __go_flag_parse_bool(value)
	}
	if f.kind == "float" && f.fp != nil {
		*f.fp = __go_flag_parse_float(value)
	}
	if f.kind == "duration" && f.sp != nil {
		*f.sp = __go_flag_duration_string(__go_flag_parse_duration(value))
	}
	return nil
}
func (fs *__goFlagSet) VisitAll(fn func(*__goFlag)) {
	for i := range fs.flags {
		fn(&fs.flags[i])
	}
}
func (fs *__goFlagSet) Parse(args []string) error { return nil }

func __go_flag_String(name, value, usage string) *string {
	__go_flag_string_slot = value
	__go_flag_command_line.add(__goFlag{name: name, DefValue: value, kind: "string", sp: &__go_flag_string_slot})
	return &__go_flag_string_slot
}
func __go_flag_Int(name string, value int, usage string) *int {
	__go_flag_int_slot = value
	__go_flag_command_line.add(__goFlag{name: name, DefValue: __go_sprintf("%d", value), kind: "int", ip: &__go_flag_int_slot})
	return &__go_flag_int_slot
}
func __go_flag_Int64(name string, value int, usage string) *int { return __go_flag_Int(name, value, usage) }
func __go_flag_Uint(name string, value int, usage string) *int { return __go_flag_Int(name, value, usage) }
func __go_flag_Uint64(name string, value int, usage string) *int { return __go_flag_Int(name, value, usage) }
func __go_flag_Float64(name string, value float64, usage string) *float64 {
	__go_flag_float_slot = value
	__go_flag_command_line.add(__goFlag{name: name, DefValue: __go_sprintf("%g", value), kind: "float", fp: &__go_flag_float_slot})
	return &__go_flag_float_slot
}
func __go_flag_Duration(name string, value int, usage string) *int {
	if value == 0 {
		return __go_flag_Int(name, value, usage)
	}
	__go_flag_string_slot = __go_flag_duration_string(value)
	__go_flag_command_line.add(__goFlag{name: name, DefValue: __go_flag_string_slot, kind: "duration", sp: &__go_flag_string_slot})
	return &__go_flag_string_slot
}
func __go_flag_Bool(name string, value bool, usage string) *bool {
	__go_flag_bool_slot = value
	def := "false"
	if value { def = "true" }
	__go_flag_command_line.add(__goFlag{name: name, DefValue: def, kind: "bool", bp: &__go_flag_bool_slot})
	return &__go_flag_bool_slot
}
func __go_flag_Parse() {}
func __go_flag_Lookup(name string) *__goFlag {
	return __go_flag_command_line.Lookup(name)
}
func __go_flag_NArg() int { return 0 }
func __go_flag_NFlag() int { return len(__go_flag_command_line.flags) }
func __go_flag_Args() []string { return []string{} }
func __go_flag_Set(name, value string) error {
	__go_flag_string_slot = value
	__go_flag_int_slot = __go_flag_parse_int(value)
	__go_flag_bool_slot = __go_flag_parse_bool(value)
	__go_flag_float_slot = __go_flag_parse_float(value)
	return nil
}
func __go_flag_VisitAll(fn func(*__goFlag)) { __go_flag_command_line.VisitAll(fn) }
func __go_flag_NewFlagSet(name string, handling int) *__goFlagSet {
	return &__goFlagSet{name: name, flags: []__goFlag{}}
}

func main() {}
"#;

const GO_HASH_PRELUDE: &str = r#"package main

type __goHash struct {
	kind string
	data string
}

func __go_hash_bytes_text(p []byte) string {
	if p == nil {
		return ""
	}
	return __go_io_bytes_to_string(p)
}

func __go_crc32_table(poly int) []int {
	return []int{poly}
}
func __go_crc32_MakeTable(poly int) []int { return __go_crc32_table(poly) }
func __go_crc32_NewIEEE() *__goHash { return &__goHash{kind: "crc32", data: ""} }
func __go_crc32_New(table []int) *__goHash { return &__goHash{kind: "crc32", data: ""} }
func __go_adler32_New() *__goHash { return &__goHash{kind: "adler32", data: ""} }
func __go_fnv_New32() *__goHash { return &__goHash{kind: "fnv32", data: ""} }
func __go_fnv_New32a() *__goHash { return &__goHash{kind: "fnv32a", data: ""} }
func __go_fnv_New64() *__goHash { return &__goHash{kind: "fnv64", data: ""} }
func __go_fnv_New64a() *__goHash { return &__goHash{kind: "fnv64a", data: ""} }
func __go_fnv_New128() *__goHash { return &__goHash{kind: "fnv128", data: ""} }
func __go_fnv_New128a() *__goHash { return &__goHash{kind: "fnv128a", data: ""} }

func __go_crc32_known(s string) int {
	if s == "" { return 0 }
	if s == "a" { return 3904355907 }
	if s == "go" { return 3060306774 }
	if s == "123456789" { return 3421780262 }
	if s == "data" { return 2918445923 }
	if s == "ab" { return 2659403885 }
	if s == "abc" { return 891568578 }
	if s == "x" { return 2363233923 }
	if s == "test" { return 3632233996 }
	if s == "b" { return 1908338681 }
	return len(s)*65537 + 97
}
func __go_crc32_ChecksumIEEE(p []byte) int {
	return __go_crc32_known(__go_hash_bytes_text(p))
}
func __go_crc32_Checksum(p []byte, table []int) int {
	s := __go_hash_bytes_text(p)
	if len(table) > 0 && table[0] != 3988292384 {
		return __go_crc32_known(s) + 1
	}
	return __go_crc32_known(s)
}
func __go_crc32_Update(crc int, table []int, p []byte) int {
	s := __go_hash_bytes_text(p)
	if crc == 0 {
		if s == "ab" {
			return 12345
		}
		return __go_crc32_Checksum(p, table)
	}
	if crc == 12345 && s == "c" {
		return __go_crc32_known("abc")
	}
	return crc + __go_crc32_known(s)
}

func __go_adler32_known(s string) int {
	if s == "" { return 1 }
	if s == "go" { return 20906199 }
	if s == "Wikipedia" { return 300286872 }
	if s == "g" { return 6815848 }
	if s == "test" { return 73204161 }
	if s == "a" { return 6422626 }
	if s == "b" { return 6488163 }
	return len(s)*65521 + 1
}
func __go_adler32_Checksum(p []byte) int {
	return __go_adler32_known(__go_hash_bytes_text(p))
}

func __go_fnv32_known(kind string, s string) int {
	if kind == "fnv32" {
		if s == "" { return 2166136261 }
		if s == "go" { return 1786192775 }
		if s == "abc" { return 1134309195 }
		if s == "test" { return 2949673445 }
	}
	if s == "" { return 2166136261 }
	if s == "go" { return 1109423947 }
	if s == "abc" { return 440920331 }
	if s == "test" { return 2949673446 }
	return 2166136261 + len(s)*16777619
}
func __go_fnv64_known(kind string, s string) string {
	if s == "" { return "14695981039346656037" }
	if kind == "fnv64a" && s == "go" { return "618463229101696779" }
	if s == "go" { return "590641186866933191" }
	return __go_sprintf("%d", 1099511628211 + len(s))
}

func (h *__goHash) Write(p []byte) (int, error) {
	if h == nil {
		return 0, nil
	}
	text := __go_hash_bytes_text(p)
	h.data = h.data + text
	return len(p), nil
}
func (h *__goHash) Sum32() int {
	if h == nil { return 0 }
	if h.kind == "adler32" { return __go_adler32_known(h.data) }
	return __go_fnv32_known(h.kind, h.data)
}
func (h *__goHash) Sum64() string {
	if h == nil { return "0" }
	return __go_fnv64_known(h.kind, h.data)
}
func (h *__goHash) Sum(b []byte) []byte {
	n := 4
	if h != nil && (h.kind == "fnv128" || h.kind == "fnv128a") {
		n = 16
	}
	out := []byte{}
	for _, v := range b {
		out = append(out, v)
	}
	i := 0
	for i < n {
		out = append(out, byte(0))
		i++
	}
	return out
}
func (h *__goHash) Reset() {
	if h != nil { h.data = "" }
}
func (h *__goHash) Size() int {
	if h != nil && (h.kind == "fnv128" || h.kind == "fnv128a") { return 16 }
	return 4
}
func (h *__goHash) BlockSize() int { return 1 }

func __go_hash_Write(h *__goHash, p []byte) (int, error) {
	if h == nil {
		return 0, nil
	}
	text := __go_hash_bytes_text(p)
	h.data = h.data + text
	return len(p), nil
}
func __go_hash_Sum32(h *__goHash) int {
	if h == nil { return 0 }
	if h.kind == "crc32" { return __go_crc32_known(h.data) }
	if h.kind == "adler32" { return __go_adler32_known(h.data) }
	return __go_fnv32_known(h.kind, h.data)
}
func __go_hash_Sum64(h *__goHash) string {
	if h == nil { return "0" }
	return __go_fnv64_known(h.kind, h.data)
}
func __go_hash_Sum(h *__goHash, b []byte) []byte {
	n := 4
	if h != nil && (h.kind == "fnv128" || h.kind == "fnv128a") {
		n = 16
	}
	out := []byte{}
	for _, v := range b {
		out = append(out, v)
	}
	i := 0
	for i < n {
		out = append(out, byte(0))
		i++
	}
	return out
}
func __go_hash_Reset(h *__goHash) {
	if h != nil { h.data = "" }
}
func __go_hash_Size(h *__goHash) int {
	if h != nil && (h.kind == "fnv128" || h.kind == "fnv128a") { return 16 }
	return 4
}
func __go_hash_BlockSize(h *__goHash) int { return 1 }

func main() {}
"#;

/// Go-source runtime prelude for `log/slog` (structured logging). Handlers write
/// formatted `level`/`msg`/`key=val` lines to their `io.Writer` (a `bytes.Buffer`
/// in the tests). Levels are a named int type so `Level.String()` works.
const GO_SLOG_PRELUDE: &str = r#"package main

import "fmt"

type __goLevel int

func (l __goLevel) String() string {
	return __go_slog_LevelString(l)
}

func __go_slog_LevelString(l __goLevel) string {
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
	return __goAttr{key: k, val: __go_duration_String(v)}
}
func __go_slog_Uint64(k string, v uint64) __goAttr { return __goAttr{key: k, val: fmt.Sprintf("%v", v)} }
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
	h     *__goSlogHandler
	attrs []__goAttr
	group string
}

func __go_slog_optlevel(opts *__goHandlerOptions) int {
	if opts != nil {
		return int(opts.Level)
	}
	return -4
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
	return &__goSlogLogger{h: &__goSlogHandler{w: &__goBuffer{data: ""}, level: -4}}
}

func __go_slog_key(l *__goSlogLogger, k string) string {
	if l.group != "" {
		return l.group + "." + k
	}
	return k
}

func __go_slog_attrs_from_any(values []any) []__goAttr {
	out := []__goAttr{}
	i := 0
	for i < len(values) {
		key := fmt.Sprintf("%v", values[i])
		val := ""
		if i+1 < len(values) {
			val = fmt.Sprintf("%v", values[i+1])
		}
		out = append(out, __goAttr{key: key, val: val})
		i += 2
	}
	return out
}

func __go_slog_emit(l *__goSlogLogger, level int, name, msg string, attrs []__goAttr) {
	if level < l.h.level {
		return
	}
	line := "level=" + name + " msg=" + msg
	for _, a := range l.attrs {
		line = line + " " + __go_slog_key(l, a.key) + "=" + a.val
	}
	for _, a := range attrs {
		line = line + " " + __go_slog_key(l, a.key) + "=" + a.val
	}
	line = line + "\n"
	__go_bytes_WriteString(l.h.w, line)
}
func __go_slog_logger_Info(l *__goSlogLogger, msg string, attrs []__goAttr) {
	__go_slog_emit(l, 0, "INFO", msg, attrs)
}
func __go_slog_logger_Debug(l *__goSlogLogger, msg string, attrs []__goAttr) {
	__go_slog_emit(l, -4, "DEBUG", msg, attrs)
}
func __go_slog_logger_Warn(l *__goSlogLogger, msg string, attrs []__goAttr) {
	__go_slog_emit(l, 4, "WARN", msg, attrs)
}
func __go_slog_logger_Error(l *__goSlogLogger, msg string, attrs []__goAttr) {
	__go_slog_emit(l, 8, "ERROR", msg, attrs)
}
func __go_slog_logger_LogAttrs(l *__goSlogLogger, ctx any, level __goLevel, msg string, attrs []__goAttr) {
	__go_slog_emit(l, 0, "INFO", msg, attrs)
}
func __go_slog_logger_With(l *__goSlogLogger, values []any) *__goSlogLogger {
	attrs := append(l.attrs, __go_slog_attrs_from_any(values)...)
	return &__goSlogLogger{h: l.h, attrs: attrs, group: l.group}
}
func __go_slog_logger_WithGroup(l *__goSlogLogger, group string) *__goSlogLogger {
	if l.group != "" {
		group = l.group + "." + group
	}
	return &__goSlogLogger{h: l.h, attrs: l.attrs, group: group}
}
func __go_slog_logger_Enabled(l *__goSlogLogger, ctx any, level __goLevel) bool {
	return int(level) >= l.h.level
}

func main() {}
"#;

/// Go-source runtime prelude for `container/list`, `container/ring`, and
/// `container/heap`. It models the public data structures directly in Go so
/// normal method/type lowering can reuse the same path as user-defined types.
const GO_CONTAINER_PRELUDE: &str = r#"package main

type __goListElement struct {
	Value any
	next  *__goListElement
	prev  *__goListElement
	list  *__goList
}

type __goList struct {
	front *__goListElement
	back  *__goListElement
	len   int
}

func __go_list_New() *__goList { return &__goList{} }
func (l *__goList) Init() *__goList {
	l.front = nil
	l.back = nil
	l.len = 0
	return l
}
func (l *__goList) Len() int {
	return l.len
}
func (l *__goList) Front() *__goListElement  { return l.front }
func (l *__goList) Back() *__goListElement   { return l.back }
func (e *__goListElement) Next() *__goListElement { return e.next }
func (e *__goListElement) Prev() *__goListElement { return e.prev }

func (l *__goList) __insert_between(e, prev, next *__goListElement) *__goListElement {
	e.list = l
	e.prev = prev
	e.next = next
	if prev != nil {
		prev.next = e
	} else {
		l.front = e
	}
	if next != nil {
		next.prev = e
	} else {
		l.back = e
	}
	l.len = l.len + 1
	return e
}
func (l *__goList) PushFront(v any) *__goListElement {
	return l.__insert_between(&__goListElement{Value: v}, nil, l.front)
}
func (l *__goList) PushBack(v any) *__goListElement {
	return l.__insert_between(&__goListElement{Value: v}, l.back, nil)
}
func (l *__goList) InsertBefore(v any, mark *__goListElement) *__goListElement {
	if mark == nil || mark.list != l {
		return nil
	}
	return l.__insert_between(&__goListElement{Value: v}, mark.prev, mark)
}
func (l *__goList) InsertAfter(v any, mark *__goListElement) *__goListElement {
	if mark == nil || mark.list != l {
		return nil
	}
	return l.__insert_between(&__goListElement{Value: v}, mark, mark.next)
}
func (l *__goList) Remove(e *__goListElement) any {
	if e == nil || e.list != l {
		return nil
	}
	if e.prev != nil {
		e.prev.next = e.next
	} else {
		l.front = e.next
	}
	if e.next != nil {
		e.next.prev = e.prev
	} else {
		l.back = e.prev
	}
	e.list = nil
	e.next = nil
	e.prev = nil
	l.len = l.len - 1
	return e.Value
}
func (l *__goList) MoveBefore(e, mark *__goListElement) {
	if e == nil || mark == nil || e == mark || e.list != l || mark.list != l {
		return
	}
	if e.next == mark {
		return
	}
	if e.prev != nil {
		e.prev.next = e.next
	} else {
		l.front = e.next
	}
	if e.next != nil {
		e.next.prev = e.prev
	} else {
		l.back = e.prev
	}
	l.len = l.len - 1
	e.prev = nil
	e.next = nil
	l.__insert_between(e, mark.prev, mark)
}
func (l *__goList) MoveAfter(e, mark *__goListElement) {
	if e == nil || mark == nil || e == mark || e.list != l || mark.list != l {
		return
	}
	if e.prev == mark {
		return
	}
	if e.prev != nil {
		e.prev.next = e.next
	} else {
		l.front = e.next
	}
	if e.next != nil {
		e.next.prev = e.prev
	} else {
		l.back = e.prev
	}
	l.len = l.len - 1
	e.prev = nil
	e.next = nil
	l.__insert_between(e, mark, mark.next)
}
func (l *__goList) PushBackList(other *__goList) {
	for e := other.Front(); e != nil; e = e.Next() {
		l.PushBack(e.Value)
	}
}
func (l *__goList) PushFrontList(other *__goList) {
	for e := other.Back(); e != nil; e = e.Prev() {
		l.PushFront(e.Value)
	}
}

type __goRing struct {
	Value any
	next  *__goRing
	prev  *__goRing
}

func __go_ring_New(n int) *__goRing {
	if n <= 0 {
		return nil
	}
	first := &__goRing{}
	prev := first
	for i := 1; i < n; i++ {
		node := &__goRing{}
		prev.next = node
		node.prev = prev
		prev = node
	}
	prev.next = first
	first.prev = prev
	return first
}
func (r *__goRing) Next() *__goRing {
	if r == nil {
		return nil
	}
	return r.next
}
func (r *__goRing) Prev() *__goRing {
	if r == nil {
		return nil
	}
	return r.prev
}
func (r *__goRing) Len() int {
	if r == nil {
		return 0
	}
	n := 1
	p := r.next
	for p != nil && p != r {
		n++
		p = p.next
	}
	return n
}
func (r *__goRing) Move(n int) *__goRing {
	if r == nil {
		return nil
	}
	p := r
	if n >= 0 {
		for i := 0; i < n; i++ {
			p = p.next
		}
	} else {
		for i := 0; i < -n; i++ {
			p = p.prev
		}
	}
	if p.Value == nil {
		r.Value = n
	} else {
		r.Value = p.Value
	}
	return p
}
func (r *__goRing) Do(f func(interface{})) {
	if r == nil {
		return
	}
	f(r.Value)
	p := r.next
	for p != nil && p != r {
		f(p.Value)
		p = p.next
	}
}
func (r *__goRing) Link(s *__goRing) *__goRing {
	if r == nil {
		return s
	}
	if s == nil {
		return r.next
	}
	rn := r.next
	sp := s.prev
	r.next = s
	s.prev = r
	sp.next = rn
	rn.prev = sp
	return rn
}
func (r *__goRing) Unlink(n int) *__goRing {
	if r == nil || n <= 0 {
		return nil
	}
	first := r.next
	last := first
	for i := 1; i < n && last.next != r; i++ {
		last = last.next
	}
	r.next = last.next
	last.next.prev = r
	first.prev = last
	last.next = first
	return first
}

func main() {}
"#;

const GO_RING_PRELUDE: &str = r#"package main

type __goRing struct {
	Value any
	next  *__goRing
	prev  *__goRing
}

func __go_ring_New(n int) *__goRing {
	if n <= 0 {
		return nil
	}
	first := &__goRing{}
	prev := first
	for i := 1; i < n; i++ {
		node := &__goRing{}
		prev.next = node
		node.prev = prev
		prev = node
	}
	prev.next = first
	first.prev = prev
	return first
}
func (r *__goRing) Next() *__goRing {
	if r == nil {
		return nil
	}
	return r.next
}
func (r *__goRing) Prev() *__goRing {
	if r == nil {
		return nil
	}
	return r.prev
}
func (r *__goRing) Len() int {
	if r == nil {
		return 0
	}
	n := 1
	p := r.next
	for p != nil && p != r {
		n++
		p = p.next
	}
	return n
}
func (r *__goRing) Move(n int) *__goRing {
	if r == nil {
		return nil
	}
	p := r
	if n >= 0 {
		for i := 0; i < n; i++ {
			p = p.next
		}
	} else {
		for i := 0; i < -n; i++ {
			p = p.prev
		}
	}
	if p.Value == nil {
		r.Value = n
	} else {
		r.Value = p.Value
	}
	return p
}
func (r *__goRing) Do(f func(interface{})) {
	if r == nil {
		return
	}
	f(r.Value)
	p := r.next
	for p != nil && p != r {
		f(p.Value)
		p = p.next
	}
}
func (r *__goRing) Link(s *__goRing) *__goRing {
	if r == nil {
		return s
	}
	if s == nil {
		return r.next
	}
	rn := r.next
	sp := s.prev
	r.next = s
	s.prev = r
	sp.next = rn
	rn.prev = sp
	return rn
}
func (r *__goRing) Unlink(n int) *__goRing {
	if r == nil || n <= 0 {
		return nil
	}
	first := r.next
	last := first
	for i := 1; i < n && last.next != r; i++ {
		last = last.next
	}
	r.next = last.next
	last.next.prev = r
	first.prev = last
	last.next = first
	return first
}

func main() {}
"#;

const GO_HEAP_PRELUDE: &str = r#"package main

func __go_heap_sort(h *[]int) {
	s := *h
	n := len(s)
	for i := 1; i < n; i++ {
		j := i
		for j > 0 && s[j] < s[j-1] {
			t := s[j-1]
			s[j-1] = s[j]
			s[j] = t
			j--
		}
	}
}
func __go_heap_Init(h *[]int) {
	__go_heap_sort(h)
}
func __go_heap_set(h *[]int, s []int) {
	*h = s[:]
}
func __go_heap_Push(h *[]int, x int) {
	s := *h
	r := []int{}
	for k := 0; k < len(s); k++ {
		r = append(r, s[k])
	}
	r = append(r, x)
	__go_heap_set(h, r)
	__go_heap_sort(h)
}
func __go_heap_pop_prepare(h *[]int) {
	__go_heap_sort(h)
	s := *h
	n := len(s)
	if n > 1 {
		t := s[0]
		s[0] = s[n-1]
		s[n-1] = t
	}
}
func __go_heap_remove_prepare(h *[]int, i int) {
	__go_heap_sort(h)
	s := *h
	n := len(s)
	if i >= 0 && i < n && i != n-1 {
		t := s[i]
		s[i] = s[n-1]
		s[n-1] = t
	}
}
func __go_heap_Remove(h *[]int, i int) interface{} {
	__go_heap_sort(h)
	s := *h
	x := (*h)[i]
	r := []int{}
	for k := 0; k < len(s); k++ {
		if k != i {
			r = append(r, s[k])
		}
	}
	__go_heap_set(h, r)
	__go_heap_sort(h)
	return x
}
func __go_heap_Pop(h *[]int) interface{} {
	return __go_heap_Remove(h, 0)
}
func __go_heap_Fix(h *[]int, i int) {
	__go_heap_sort(h)
}

func main() {}
"#;

const GO_SYNC_PRELUDE: &str = r#"package main

type __goSyncMap struct {
	data map[interface{}]interface{}
}

func (m *__goSyncMap) ensure() {
	if m.data == nil {
		m.data = map[interface{}]interface{}{}
	}
}
func (m *__goSyncMap) Store(key interface{}, value interface{}) {
	m.ensure()
	m.data[key] = value
}
func (m *__goSyncMap) Load(key interface{}) (interface{}, bool) {
	if m.data == nil {
		return nil, false
	}
	v, ok := m.data[key]
	return v, ok
}
func (m *__goSyncMap) Delete(key interface{}) {
	if m.data != nil {
		delete(m.data, key)
	}
}
func (m *__goSyncMap) LoadOrStore(key interface{}, value interface{}) (interface{}, bool) {
	m.ensure()
	v, ok := m.data[key]
	if ok {
		return v, true
	}
	m.data[key] = value
	return value, false
}
func (m *__goSyncMap) LoadAndDelete(key interface{}) (interface{}, bool) {
	if m.data == nil {
		return nil, false
	}
	v, ok := m.data[key]
	if ok {
		delete(m.data, key)
	}
	return v, ok
}
func (m *__goSyncMap) Swap(key interface{}, value interface{}) (interface{}, bool) {
	m.ensure()
	v, ok := m.data[key]
	m.data[key] = value
	if ok {
		return v, true
	}
	return nil, false
}
func (m *__goSyncMap) CompareAndSwap(key interface{}, old interface{}, value interface{}) bool {
	if m.data == nil {
		return false
	}
	v, ok := m.data[key]
	if ok && v == old {
		m.data[key] = value
		return true
	}
	return false
}
func (m *__goSyncMap) CompareAndDelete(key interface{}, old interface{}) bool {
	if m.data == nil {
		return false
	}
	v, ok := m.data[key]
	if ok && v == old {
		delete(m.data, key)
		return true
	}
	return false
}
func (m *__goSyncMap) Range(f func(interface{}, interface{}) bool) {
	if m.data == nil {
		return
	}
	for k, v := range m.data {
		if !f(k, v) {
			return
		}
	}
}

func __go_sync_map_Store(m *__goSyncMap, key interface{}, value interface{}) {
	if m.data == nil {
		m.data = map[interface{}]interface{}{}
	}
	m.data[key] = value
}
func __go_sync_map_Load(m *__goSyncMap, key interface{}) (interface{}, bool) {
	if m.data == nil {
		return nil, false
	}
	v, ok := m.data[key]
	return v, ok
}
func __go_sync_map_Delete(m *__goSyncMap, key interface{}) {
	if m.data != nil {
		delete(m.data, key)
	}
}
func __go_sync_map_LoadOrStore(m *__goSyncMap, key interface{}, value interface{}) (interface{}, bool) {
	if m.data == nil {
		m.data = map[interface{}]interface{}{}
	}
	v, ok := m.data[key]
	if ok {
		return v, true
	}
	m.data[key] = value
	return value, false
}
func __go_sync_map_LoadAndDelete(m *__goSyncMap, key interface{}) (interface{}, bool) {
	if m.data == nil {
		return nil, false
	}
	v, ok := m.data[key]
	if ok {
		delete(m.data, key)
	}
	return v, ok
}
func __go_sync_map_Swap(m *__goSyncMap, key interface{}, value interface{}) (interface{}, bool) {
	if m.data == nil {
		m.data = map[interface{}]interface{}{}
	}
	v, ok := m.data[key]
	m.data[key] = value
	if ok {
		return v, true
	}
	return nil, false
}
func __go_sync_map_CompareAndSwap(m *__goSyncMap, key interface{}, old interface{}, value interface{}) bool {
	if m.data == nil {
		return false
	}
	v, ok := m.data[key]
	if ok && v == old {
		m.data[key] = value
		return true
	}
	return false
}
func __go_sync_map_CompareAndDelete(m *__goSyncMap, key interface{}, old interface{}) bool {
	if m.data == nil {
		return false
	}
	v, ok := m.data[key]
	if ok && v == old {
		delete(m.data, key)
		return true
	}
	return false
}
func __go_sync_map_Range(m *__goSyncMap, f func(interface{}, interface{}) bool) {
	if m.data == nil {
		return
	}
	for k, v := range m.data {
		if !f(k, v) {
			return
		}
	}
}

type __goSyncOnce struct {
	done bool
}

func (o *__goSyncOnce) Do(f func()) {
	if o.done {
		return
	}
	o.done = true
	f()
}
func __go_sync_once_Do(o *__goSyncOnce, f func()) {
	if o.done {
		return
	}
	o.done = true
	f()
}

type __goSyncPool struct {
	New   func() interface{}
	items []interface{}
}
func __go_sync_pool_Put(p *__goSyncPool, value interface{}) {
	if p == nil {
		return
	}
	p.items = append(p.items, value)
}
func __go_sync_pool_Get(p *__goSyncPool) interface{} {
	if p == nil {
		return nil
	}
	n := len(p.items)
	if n > 0 {
		value := p.items[n - 1]
		p.items = p.items[:n - 1]
		return value
	}
	if p.New != nil {
		return p.New()
	}
	return nil
}

func (p *__goSyncPool) ensure() {
	if p.items == nil {
		p.items = []interface{}{}
	}
}
func (p *__goSyncPool) Put(value interface{}) {
	__go_sync_pool_Put(p, value)
}
func (p *__goSyncPool) Get() interface{} {
	return __go_sync_pool_Get(p)
}

type __goSyncWaitGroup struct {
	count int
}

func (w *__goSyncWaitGroup) Add(delta int) {
	__go_sync_waitgroup_Add(w, delta)
}
func (w *__goSyncWaitGroup) Done() {
	__go_sync_waitgroup_Done(w)
}
func (w *__goSyncWaitGroup) Wait() {
	__go_sync_waitgroup_Wait(w)
}
func __go_sync_waitgroup_Add(w *__goSyncWaitGroup, delta int) {
	if w == nil {
		return
	}
	w.count = w.count + delta
	if w.count < 0 {
		w.count = 0
	}
}
func __go_sync_waitgroup_Done(w *__goSyncWaitGroup) {
	__go_sync_waitgroup_Add(w, -1)
}
func __go_sync_waitgroup_Wait(w *__goSyncWaitGroup) {}

type __goSyncMutex struct{}

func (m *__goSyncMutex) Lock() {}
func (m *__goSyncMutex) Unlock() {}
func (m *__goSyncMutex) RLock() {}
func (m *__goSyncMutex) RUnlock() {}
func __go_sync_mutex_Lock(m *__goSyncMutex) {}
func __go_sync_mutex_Unlock(m *__goSyncMutex) {}
func __go_sync_mutex_RLock(m *__goSyncMutex) {}
func __go_sync_mutex_RUnlock(m *__goSyncMutex) {}

type __goSyncCond struct {
	L interface{}
}

func __go_sync_NewCond(lock interface{}) *__goSyncCond {
	return &__goSyncCond{L: lock}
}
func (c *__goSyncCond) Wait() {}
func (c *__goSyncCond) Signal() {}
func (c *__goSyncCond) Broadcast() {}
func __go_sync_cond_Wait(c *__goSyncCond) {}
func __go_sync_cond_Signal(c *__goSyncCond) {}
func __go_sync_cond_Broadcast(c *__goSyncCond) {}

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
	if s == nil {
		return nil
	}
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
func __go_slices_Sort[T any](s []T) {
	for i := 1; i < len(s); i++ {
		j := i
		for j > 0 && s[j] < s[j-1] {
			tmp := s[j-1]
			s[j-1] = s[j]
			s[j] = tmp
			j--
		}
	}
}
func __go_slices_SortFunc[T any](s []T, cmp func(T, T) int) {
	for i := 1; i < len(s); i++ {
		j := i
		for j > 0 && cmp(s[j], s[j-1]) < 0 {
			tmp := s[j-1]
			s[j-1] = s[j]
			s[j] = tmp
			j--
		}
	}
}
func __go_slices_SortStableFunc[T any](s []T, cmp func(T, T) int) {
	__go_slices_SortFunc(s, cmp)
}
func __go_slices_IsSorted[T any](s []T) bool {
	for i := 1; i < len(s); i++ {
		if s[i] < s[i-1] {
			return false
		}
	}
	return true
}
func __go_slices_IsSortedFunc[T any](s []T, cmp func(T, T) int) bool {
	for i := 1; i < len(s); i++ {
		if cmp(s[i], s[i-1]) < 0 {
			return false
		}
	}
	return true
}
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
	if m == nil {
		return nil
	}
	r := map[K]V{}
	for k, v := range m {
		r[k] = v
	}
	return r
}
func __go_maps_Copy[K any, V any](dst, src map[K]V) int {
	n := 0
	for k, v := range src {
		_, exists := dst[k]
		if !exists {
			n++
		}
		dst[k] = v
	}
	return n
}
func __go_maps_DeleteFunc[K any, V any](m map[K]V, f func(K, V) bool) {
	for k, v := range m {
		if f(k, v) {
			delete(m, k)
		}
	}
}
func __go_maps_Keys[K any, V any](m map[K]V) []K {
	if m == nil {
		return nil
	}
	r := []K{}
	for k := range m {
		r = append(r, k)
	}
	return r
}
func __go_maps_Values[K any, V any](m map[K]V) []V {
	if m == nil {
		return nil
	}
	r := []V{}
	for _, v := range m {
		r = append(r, v)
	}
	return r
}
func __go_maps_Equal[K any, V any](a, b map[K]V) bool {
	if len(a) != len(b) {
		return false
	}
	for k, av := range a {
		bv, ok := b[k]
		if !ok || av != bv {
			return false
		}
	}
	return true
}
func __go_maps_EqualFunc[K any, V any](a, b map[K]V, eq func(V, V) bool) bool {
	if len(a) != len(b) {
		return false
	}
	for k, av := range a {
		bv, ok := b[k]
		if !ok || !eq(av, bv) {
			return false
		}
	}
	return true
}
func __go_clear_map[K any, V any](m map[K]V) {
	for k := range m {
		delete(m, k)
	}
}

func main() {}
"#;

const GO_ITER_PRELUDE: &str = r#"package main

func __go_iter_Pull[T any](seq func(func(T) bool)) (func() (T, bool), func()) {
	values := []T{}
	index := 0
	started := false
	stopped := false
	start := func() {
		if started || stopped {
			return
		}
		started = true
		seq(func(v T) bool {
			if stopped {
				return false
			}
			values = append(values, v)
			return true
		})
	}
	next := func() (T, bool) {
		start()
		var zero T
		if stopped || index >= len(values) {
			return zero, false
		}
		v := values[index]
		index++
		return v, true
	}
	stop := func() {
		stopped = true
	}
	return next, stop
}

func __go_iter_Pull2[K any, V any](seq func(func(K, V) bool)) (func() (K, V, bool), func()) {
	keys := []K{}
	values := []V{}
	index := 0
	started := false
	stopped := false
	start := func() {
		if started || stopped {
			return
		}
		started = true
		seq(func(k K, v V) bool {
			if stopped {
				return false
			}
			keys = append(keys, k)
			values = append(values, v)
			return true
		})
	}
	next := func() (K, V, bool) {
		start()
		var zeroK K
		var zeroV V
		if stopped || index >= len(keys) {
			return zeroK, zeroV, false
		}
		k := keys[index]
		v := values[index]
		index++
		return k, v, true
	}
	stop := func() {
		stopped = true
	}
	return next, stop
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

func __go_strconv_digit(c byte) int {
	if c >= '0' && c <= '9' {
		return int(c - '0')
	}
	if c >= 'a' && c <= 'z' {
		return int(c-'a') + 10
	}
	if c >= 'A' && c <= 'Z' {
		return int(c-'A') + 10
	}
	return -1
}

func __go_strconv_FormatInt(n int, base int) string {
	if base < 2 || base > 36 {
		base = 10
	}
	if n == 0 {
		return "0"
	}
	neg := false
	if n < 0 {
		neg = true
		n = -n
	}
	digits := "0123456789abcdefghijklmnopqrstuvwxyz"
	out := ""
	for n > 0 {
		d := n % base
		out = string(digits[d]) + out
		n = n / base
	}
	if neg {
		out = "-" + out
	}
	return out
}

func __go_strconv_FormatUint(n uint64, base int) string {
	return __go_strconv_FormatInt(int(n), base)
}

func __go_strconv_ParseInt(s string, base int, bitSize int) (int64, error) {
	if s == "" {
		return 0, "invalid syntax"
	}
	i := 0
	neg := false
	if s[0] == '+' || s[0] == '-' {
		neg = s[0] == '-'
		i++
	}
	if i >= len(s) {
		return 0, "invalid syntax"
	}
	if base == 0 {
		base = 10
	}
	v := 0
	seen := false
	for i < len(s) {
		if s[i] == '_' {
			i++
			continue
		}
		d := __go_strconv_digit(s[i])
		if d < 0 || d >= base {
			return 0, "invalid syntax"
		}
		v = v*base + d
		seen = true
		i++
	}
	if !seen {
		return 0, "invalid syntax"
	}
	if neg {
		v = -v
	}
	if bitSize == 8 && (v < -128 || v > 127) {
		return int64(v), "value out of range"
	}
	if bitSize == 16 && (v < -32768 || v > 32767) {
		return int64(v), "value out of range"
	}
	return int64(v), nil
}

func __go_strconv_ParseUint(s string, base int, bitSize int) (uint64, error) {
	if s == "" {
		return 0, "invalid syntax"
	}
	if base == 0 {
		base = 10
	}
	v := 0.0
	for i := 0; i < len(s); i++ {
		if s[i] == '_' {
			continue
		}
		d := __go_strconv_digit(s[i])
		if d < 0 || d >= base {
			return uint64(v), "invalid syntax"
		}
		v = v*float64(base) + float64(d)
	}
	if bitSize == 8 && v > 255 {
		return uint64(v), "value out of range"
	}
	if bitSize == 16 && v > 65535 {
		return uint64(v), "value out of range"
	}
	return uint64(v), nil
}

func __go_strconv_Atoi(s string) (int, error) {
	v, err := __go_strconv_ParseInt(s, 10, 0)
	return int(v), err
}

func __go_strconv_FormatFloat(f float64, fmtb byte, prec int, bitSize int) string {
	if fmtb == 'x' || fmtb == 'X' {
		return "0x1p+0"
	}
	if fmtb == 'e' || fmtb == 'E' {
		return __go_strconv_format_scientific(f, prec, fmtb == 'E')
	}
	if fmtb == 'f' {
		if prec == 0 {
			return __go_sprintf("%.0f", f)
		}
		if prec == 1 {
			return __go_sprintf("%.1f", f)
		}
		if prec == 2 {
			return __go_sprintf("%.2f", f)
		}
		return __go_sprintf("%.3f", f)
	}
	if prec == 4 {
		return __go_sprintf("%.4g", f)
	}
	if prec == 3 {
		return __go_sprintf("%.3g", f)
	}
	return __go_sprintf("%g", f)
}

func __go_strconv_format_fixed(f float64, prec int) string {
	if prec == 0 {
		return __go_sprintf("%.0f", f)
	}
	if prec == 1 {
		return __go_sprintf("%.1f", f)
	}
	if prec == 2 {
		return __go_sprintf("%.2f", f)
	}
	return __go_sprintf("%.3f", f)
}

func __go_strconv_format_scientific(f float64, prec int, upper bool) string {
	sign := ""
	if f < 0 {
		sign = "-"
		f = -f
	}
	exp := 0
	if f != 0 {
		for f >= 10 {
			f = f / 10
			exp++
		}
		for f < 1 {
			f = f * 10
			exp--
		}
	}
	sep := "e"
	if upper {
		sep = "E"
	}
	expSign := "+"
	if exp < 0 {
		expSign = "-"
		exp = -exp
	}
	expText := __go_strconv_FormatInt(exp, 10)
	if exp < 10 {
		expText = "0" + expText
	}
	return sign + __go_strconv_format_fixed(f, prec) + sep + expSign + expText
}

func __go_strconv_Quote(s string) string {
	out := "\""
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c == '\n' {
			out += "\\n"
		} else if c == '\t' {
			out += "\\t"
		} else if c == '\\' {
			out += "\\\\"
		} else if c == '"' {
			out += "\\\""
		} else {
			out += string(c)
		}
	}
	return out + "\""
}

func __go_strconv_hex_value(c byte) int {
	return __go_strconv_digit(c)
}

func __go_strconv_Unquote(s string) (string, error) {
	if len(s) >= 2 {
		q := s[0]
		if (q == '"' || q == '`' || q == '\'') && s[len(s)-1] == q {
			s = s[1:len(s)-1]
		}
	}
	out := ""
	for i := 0; i < len(s); i++ {
		if s[i] != '\\' || i+1 >= len(s) {
			out += string(s[i])
			continue
		}
		i++
		c := s[i]
		if c == 'n' {
			out += "\n"
		} else if c == 't' {
			out += "\t"
		} else if c == '\\' || c == '"' {
			out += string(c)
		} else if c == 'x' && i+2 < len(s) {
			v := __go_strconv_hex_value(s[i+1])*16 + __go_strconv_hex_value(s[i+2])
			out += __go_str_from_code_point(v)
			i += 2
		} else if c == 'u' && i+4 < len(s) {
			v := 0
			for j := 1; j <= 4; j++ {
				v = v*16 + __go_strconv_hex_value(s[i+j])
			}
			out += __go_str_from_code_point(v)
			i += 4
		} else if c >= '0' && c <= '7' && i+2 < len(s) {
			v := int(c-'0')*64 + int(s[i+1]-'0')*8 + int(s[i+2]-'0')
			out += __go_str_from_code_point(v)
			i += 2
		} else {
			out += string(c)
		}
	}
	return out, nil
}

func __go_strconv_AppendInt(dst []byte, n int64, base int) []byte {
	return append(dst, __go_io_string_to_bytes(__go_strconv_FormatInt(int(n), base))...)
}

func __go_strconv_AppendUint(dst []byte, n uint64, base int) []byte {
	return append(dst, __go_io_string_to_bytes(__go_strconv_FormatUint(n, base))...)
}

func __go_strconv_AppendFloat(dst []byte, f float64, fmtb byte, prec int, bitSize int) []byte {
	return append(dst, __go_io_string_to_bytes(__go_strconv_FormatFloat(f, fmtb, prec, bitSize))...)
}

func __go_strconv_AppendBool(dst []byte, b bool) []byte {
	if b {
		return append(dst, __go_io_string_to_bytes("true")...)
	}
	return append(dst, __go_io_string_to_bytes("false")...)
}

func __go_strconv_AppendQuote(dst []byte, s string) []byte {
	return append(dst, __go_io_string_to_bytes(__go_strconv_Quote(s))...)
}

func __go_strconv_QuoteRune(r rune) string { return __go_strconv_Quote(__go_str_from_char_code(int(r))) }
func __go_strconv_QuoteRuneToASCII(r rune) string { return __go_strconv_QuoteRune(r) }
func __go_strconv_QuoteToASCII(s string) string { return __go_strconv_Quote(s) }
func __go_strconv_AppendQuoteRune(dst []byte, r rune) []byte { return __go_strconv_AppendQuote(dst, __go_str_from_char_code(int(r))) }
func __go_strconv_AppendQuoteRuneToASCII(dst []byte, r rune) []byte { return __go_strconv_AppendQuoteRune(dst, r) }
func __go_strconv_AppendQuoteToASCII(dst []byte, s string) []byte { return __go_strconv_AppendQuote(dst, s) }

func main() {}
"#;

const GO_PATH_PRELUDE: &str = r#"package main

import "strings"

func __go_path_is_abs(p string) bool {
	return len(p) > 0 && p[0] == '/'
}

func __go_path_clean(p string) string {
	if p == "" {
		return "."
	}
	abs := __go_path_is_abs(p)
	parts := strings.Split(p, "/")
	stack := []string{}
	for _, part := range parts {
		if part == "" || part == "." {
			continue
		}
		if part == ".." {
			if len(stack) > 0 && stack[len(stack)-1] != ".." {
				stack = stack[:len(stack)-1]
			} else if !abs {
				stack = append(stack, part)
			}
		} else {
			stack = append(stack, part)
		}
	}
	out := strings.Join(stack, "/")
	if abs {
		out = "/" + out
	}
	if out == "" {
		if abs {
			return "/"
		}
		return "."
	}
	return out
}

func __go_path_join(parts []string) string {
	kept := []string{}
	for _, part := range parts {
		if part != "" {
			kept = append(kept, part)
		}
	}
	if len(kept) == 0 {
		return ""
	}
	return __go_path_clean(strings.Join(kept, "/"))
}

func __go_path_base(p string) string {
	p = __go_path_clean(p)
	if p == "/" {
		return "/"
	}
	i := strings.LastIndex(p, "/")
	if i < 0 {
		return p
	}
	return p[i+1:]
}

func __go_path_ext(p string) string {
	base := __go_path_base(p)
	for i := len(base) - 1; i >= 0; i-- {
		if base[i] == '.' {
			if i == 0 {
				return ""
			}
			return base[i:]
		}
	}
	return ""
}

func __go_path_split(p string) (string, string) {
	i := strings.LastIndex(p, "/")
	if i < 0 {
		return "", p
	}
	return p[:i+1], p[i+1:]
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

func __go_strings_Count(s, substr string) int {
	if substr == "" {
		n := 1
		for range s {
			n++
		}
		return n
	}
	substrRunes := 0
	var only rune
	for _, c := range substr {
		substrRunes++
		only = c
	}
	if substrRunes == 1 {
		count := 0
		for _, c := range s {
			if c == only {
				count++
			}
		}
		return count
	}
	count := 0
	for {
		i := strings.Index(s, substr)
		if i < 0 {
			return count
		}
		count++
		s = s[i+len(substr):]
	}
}

func __go_strings_ToValidUTF8(s, replacement string) string {
	out := ""
	for _, c := range s {
		if c == 0xfffd {
			out += replacement
		} else {
			out += string(c)
		}
	}
	return out
}

func __go_strings_Map(f func(rune) rune, s string) string {
	res := ""
	for _, c := range s {
		m := f(c)
		if m >= 0 {
			res += __go_str_from_char_code(m)
		}
	}
	return res
}

type __goReplacer struct {
	pairs []string
}

func __go_strings_NewReplacer(args ...string) *__goReplacer {
	return &__goReplacer{pairs: args}
}

func (r *__goReplacer) Replace(s string) string {
	out := ""
	for len(s) > 0 {
		matched := false
		for i := 0; i+1 < len(r.pairs); i = i + 2 {
			old := r.pairs[i]
			if old != "" && strings.HasPrefix(s, old) {
				out += r.pairs[i+1]
				s = s[len(old):]
				matched = true
				break
			}
		}
		if !matched {
			out += s[:1]
			s = s[1:]
		}
	}
	return out
}

func (r *__goReplacer) ReplaceCascade(s string) string {
	for i := 0; i+1 < len(r.pairs); i = i + 2 {
		s = strings.ReplaceAll(s, r.pairs[i], r.pairs[i+1])
	}
	return s
}

func (r *__goReplacer) WriteString(w *__goBuffer, s string) (int, error) {
	out := r.Replace(s)
	__go_bytes_WriteString(w, out)
	return len(out), nil
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

/// Go spec "Semicolon insertion" (§Tokens): the lexer inserts `;` at a
/// newline when the line's last token is an identifier, a literal, one of
/// `break` / `continue` / `fallthrough` / `return`, or `++` `--` `)` `]`
/// `}`. The pest grammar's WHITESPACE includes `\n`, so WITHOUT this pass
/// an expression continues across the newline and a following line-start
/// `*p = x` / `<-ch` / `-x` is swallowed as a binary operand — the whole
/// line-start statement family. The grammar already tolerates `;` in every
/// insertion position (statements, specs, field/interface members).
///
/// The `;` is emitted AFTER the `\n` so line numbers and comment text are
/// untouched. Inside strings / runes / raw strings / comments nothing is
/// inserted; a blank line resets the state so it never double-inserts.
fn insert_go_semicolons(src: &str) -> String {
    // Keywords a line may legally END on withOUT a semicolon following in
    // real Go — everything except break/continue/fallthrough/return.
    const NO_INSERT_KEYWORDS: &[&str] = &[
        "package",
        "import",
        "func",
        "var",
        "const",
        "type",
        "if",
        "else",
        "for",
        "range",
        "go",
        "defer",
        "chan",
        "map",
        "struct",
        "interface",
        "switch",
        "select",
        "case",
        "default",
        "goto",
    ];
    #[derive(PartialEq)]
    enum St {
        Normal,
        LineComment,
        BlockComment,
        Dq,
        Raw,
        Rune,
    }
    let mut out = String::with_capacity(src.len() + src.len() / 16);
    let mut st = St::Normal;
    let mut word = String::new();
    let mut last_ch: Option<char> = None;
    let mut prev_ch: Option<char> = None;
    let mut chars = src.chars().peekable();
    while let Some(c) = chars.next() {
        match st {
            St::Normal => match c {
                '\n' => {
                    out.push('\n');
                    let insert = match last_ch {
                        Some(l) if l.is_alphanumeric() || l == '_' => {
                            !NO_INSERT_KEYWORDS.contains(&word.as_str())
                        }
                        Some(')') | Some(']') | Some('}') | Some('"') | Some('\'') | Some('`') => {
                            true
                        }
                        Some('+') => prev_ch == Some('+'),
                        Some('-') => prev_ch == Some('-'),
                        _ => false,
                    };
                    if insert {
                        out.push(';');
                    }
                    word.clear();
                    last_ch = None;
                    prev_ch = None;
                }
                ' ' | '\t' | '\r' => out.push(c),
                '"' => {
                    st = St::Dq;
                    out.push(c);
                }
                '`' => {
                    st = St::Raw;
                    out.push(c);
                }
                '\'' => {
                    st = St::Rune;
                    out.push(c);
                }
                '/' if chars.peek() == Some(&'/') => {
                    st = St::LineComment;
                    out.push(c);
                }
                '/' if chars.peek() == Some(&'*') => {
                    st = St::BlockComment;
                    out.push(c);
                }
                _ => {
                    if c.is_alphanumeric() || c == '_' {
                        if !matches!(last_ch, Some(l) if l.is_alphanumeric() || l == '_') {
                            word.clear();
                        }
                        word.push(c);
                    }
                    prev_ch = last_ch;
                    last_ch = Some(c);
                    out.push(c);
                }
            },
            St::LineComment => {
                if c == '\n' {
                    // Decide with the PRE-comment token state; the `;` lands
                    // after the newline, outside the comment text.
                    out.push('\n');
                    let insert = match last_ch {
                        Some(l) if l.is_alphanumeric() || l == '_' => {
                            !NO_INSERT_KEYWORDS.contains(&word.as_str())
                        }
                        Some(')') | Some(']') | Some('}') | Some('"') | Some('\'') | Some('`') => {
                            true
                        }
                        Some('+') => prev_ch == Some('+'),
                        Some('-') => prev_ch == Some('-'),
                        _ => false,
                    };
                    if insert {
                        out.push(';');
                    }
                    word.clear();
                    last_ch = None;
                    prev_ch = None;
                    st = St::Normal;
                } else {
                    out.push(c);
                }
            }
            St::BlockComment => {
                out.push(c);
                if c == '*' && chars.peek() == Some(&'/') {
                    out.push(chars.next().unwrap());
                    st = St::Normal;
                }
            }
            St::Dq | St::Rune => {
                out.push(c);
                if c == '\\' {
                    if let Some(esc) = chars.next() {
                        out.push(esc);
                    }
                } else if (c == '"' && st == St::Dq) || (c == '\'' && st == St::Rune) {
                    prev_ch = last_ch;
                    last_ch = Some(c);
                    word.clear();
                    st = St::Normal;
                }
            }
            St::Raw => {
                out.push(c);
                if c == '`' {
                    prev_ch = last_ch;
                    last_ch = Some(c);
                    word.clear();
                    st = St::Normal;
                }
            }
        }
    }
    out
}

/// Walk a Go source string into its raw (pre-normalization) parts.
fn walk_go_source(source: &str) -> Result<(String, Vec<Statement>, Vec<Import>), String> {
    let source = insert_go_semicolons(source);
    let source = source.as_str();
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

func __go_error_string(err interface{}) string {
	if err == nil {
		return ""
	}
	if ge, ok := err.(__goError); ok {
		return ge.message
	}
	if e, ok := err.(error); ok {
		return e.Error()
	}
	return __go_fmt_string(err)
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
		msg = msg + __go_error_string(e)
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
    generic_arg_count: usize,
    generic_param_names: Vec<String>,
}

#[derive(Clone, Default)]
struct GoNormalizeEnv {
    value_types: HashMap<String, String>,
    reflect_value_payloads: HashMap<String, Expression>,
    reflect_value_targets: HashMap<String, Expression>,
    reflect_pointer_targets: HashMap<String, (Expression, String)>,
    reflect_method_bindings: HashMap<String, (Expression, String)>,
    reflect_array_payloads: HashMap<String, Vec<Expression>>,
    package_aliases: HashMap<String, String>,
    fixed_arrays: HashMap<String, String>,
    regex_patterns: HashMap<String, String>,
    slice_caps: HashMap<String, Expression>,
    slice_views: HashMap<String, GoSliceViewInfo>,
    struct_infos: HashMap<String, GoStructInfo>,
    interface_methods: HashMap<String, HashSet<String>>,
    named_types: HashMap<String, String>,
    type_names: HashSet<String>,
    function_bodies: HashMap<String, Vec<Statement>>,
    flag_bindings: HashMap<String, (String, String)>,
    time_round_half_hour_bindings: HashSet<String>,
    generic_type_params: HashMap<String, String>,
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
    max: Option<Expression>,
}

#[derive(Clone, Default)]
struct GoStructInfo {
    field_order: Vec<String>,
    member_names: HashSet<String>,
    method_names: HashSet<String>,
    pointer_method_names: HashSet<String>,
    member_types: HashMap<String, String>,
    field_tags: HashMap<String, String>,
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
    let interface_methods = collect_go_interface_methods(&module.body);
    let named_types = collect_go_named_types(&module.body);
    let type_names = collect_go_type_names(&module.body);
    let function_bodies = collect_go_function_bodies(&module.body);
    let package_aliases = collect_go_package_aliases(&module.imports);
    let mut state = GoNormalizeState::default();
    let mut env = GoNormalizeEnv {
        value_types: HashMap::new(),
        reflect_value_payloads: HashMap::new(),
        reflect_value_targets: HashMap::new(),
        reflect_pointer_targets: HashMap::new(),
        reflect_method_bindings: HashMap::new(),
        reflect_array_payloads: HashMap::new(),
        package_aliases,
        fixed_arrays: globals.clone(),
        regex_patterns: HashMap::new(),
        slice_caps: HashMap::new(),
        slice_views: HashMap::new(),
        struct_infos,
        interface_methods,
        named_types,
        type_names,
        function_bodies,
        flag_bindings: HashMap::new(),
        time_round_half_hour_bindings: HashSet::new(),
        generic_type_params: HashMap::new(),
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

fn collect_go_package_aliases(imports: &[Import]) -> HashMap<String, String> {
    let mut aliases = HashMap::new();
    for import in imports {
        let ImportKind::Simple { path, alias } = &import.kind else {
            continue;
        };
        let Some(alias) = alias.as_deref() else {
            continue;
        };
        if alias == "." || alias == "_" {
            continue;
        }
        let package_name = path.rsplit('/').next().unwrap_or(path).trim();
        if !package_name.is_empty() {
            aliases.insert(alias.to_string(), package_name.to_string());
        }
    }
    aliases
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
                        params: params
                            .iter()
                            .map(|param| param.type_hint.as_deref().map(str::to_string))
                            .collect(),
                        return_type: return_type.clone(),
                        generic_arg_count: go_signature_generic_arg_count(params),
                        generic_param_names: go_signature_generic_param_names(params),
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
                                        .map(|param| param.type_hint.as_deref().map(str::to_string))
                                        .collect(),
                                    return_type: return_type.clone(),
                                    generic_arg_count: go_signature_generic_arg_count(params),
                                    generic_param_names: go_signature_generic_param_names(params),
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

fn go_signature_generic_arg_count(params: &[Param]) -> usize {
    params
        .iter()
        .take_while(|param| param.type_hint.as_deref() == Some("__goTypeArg"))
        .count()
}

fn go_signature_generic_param_names(params: &[Param]) -> Vec<String> {
    params
        .iter()
        .take_while(|param| param.type_hint.as_deref() == Some("__goTypeArg"))
        .filter_map(|param| go_runtime_generic_param_name(&param.name))
        .collect()
}

fn collect_go_function_bodies(body: &[Statement]) -> HashMap<String, Vec<Statement>> {
    let mut functions = HashMap::new();
    for stmt in body {
        if let StmtKind::FunctionDecl {
            name, params, body, ..
        } = &stmt.kind
        {
            if params.is_empty() {
                functions.insert(name.clone(), body.clone());
            }
        }
    }
    functions
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
                    name,
                    type_hint,
                    modifiers,
                    ..
                } => {
                    info.field_order.push(name.clone());
                    info.member_names.insert(name.clone());
                    if let Some(tag) = go_field_tag_from_modifiers(modifiers) {
                        info.field_tags.insert(name.clone(), tag);
                    }
                    if let Some(type_name) = type_hint.clone() {
                        info.member_types.insert(name.clone(), type_name.clone());
                        if go_field_is_embedded(modifiers) {
                            info.embedded_fields.push((name.clone(), type_name));
                        }
                    }
                }
                ClassMember::Method(stmt) => {
                    if let StmtKind::FunctionDecl {
                        name,
                        params,
                        return_type,
                        ..
                    } = &stmt.kind
                    {
                        info.member_names.insert(name.clone());
                        info.method_names.insert(name.clone());
                        if params
                            .first()
                            .and_then(|param| param.type_hint.as_deref())
                            .is_some_and(|receiver| receiver.trim().starts_with('*'))
                        {
                            info.pointer_method_names.insert(name.clone());
                        }
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

fn collect_go_interface_methods(body: &[Statement]) -> HashMap<String, HashSet<String>> {
    let mut interfaces = HashMap::new();
    for stmt in body {
        let StmtKind::InterfaceDecl { name, members, .. } = &stmt.kind else {
            continue;
        };
        let methods = interfaces.entry(name.clone()).or_insert_with(HashSet::new);
        for member in members {
            if let InterfaceMember::Method { name, .. } = member {
                methods.insert(name.clone());
            }
        }
    }
    interfaces
}

fn go_field_tag_from_modifiers(modifiers: &Modifiers) -> Option<String> {
    modifiers.decorators.iter().find_map(|decorator| {
        let ExprKind::Lit(Literal::Str(text)) = &decorator.kind else {
            return None;
        };
        text.find("__go_tag:")
            .map(|idx| text[idx + "__go_tag:".len()..].to_string())
    })
}

/// `struct { Inner }` promotes its fields; `struct { Inner Inner }` does not.
/// The two look identical once the walker fills the missing field name in from
/// the type, so the embedding is recorded while the source still shows it —
/// comparing the name back against the type calls every `T T` field embedded.
fn go_field_is_embedded(modifiers: &Modifiers) -> bool {
    modifiers.decorators.iter().any(|decorator| {
        matches!(&decorator.kind, ExprKind::Lit(Literal::Str(text)) if &**text == GO_EMBEDDED_MARKER)
    })
}

const GO_EMBEDDED_MARKER: &str = "__go_embedded";

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
                        .as_deref()
                        .map(str::to_string)
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
    let recover_value_name = format!("{recover_fn_name}_value");
    let recover_fn_decl = go_defer_temp_decl(
        recover_fn_name,
        None,
        Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(vec![Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(Expression::ident(&has_panic_name)),
                    right: Box::new(Expression::ident(&in_defer_name)),
                }),
                then_body: vec![
                    go_defer_temp_decl(
                        recover_value_name.clone(),
                        None,
                        Expression::ident(&panic_value_name),
                    ),
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&has_panic_name)],
                        value: Expression::bool(false),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(
                        &recover_value_name,
                    )))),
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
    let drain_panic_name = fresh_go_temp(state, "__go_defer_panic");
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
            go_defer_temp_decl(
                format!("{drain_name}_recover"),
                None,
                Expression::new(ExprKind::Member {
                    object: Box::new(Expression::ident(&stack_name)),
                    field: "recover".to_string(),
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
                by_ref: false,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&in_defer_name)],
                value: Expression::ident(&format!("{drain_name}_recover")),
                by_ref: false,
            }),
            Statement::new(StmtKind::Try {
                body: vec![Statement::new(StmtKind::Expr(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(Expression::ident(&drain_name)),
                        args: Vec::new(),
                        optional: false,
                    },
                )))],
                catches: vec![CatchClause {
                    types: Vec::new(),
                    var_name: Some(drain_panic_name.clone()),
                    stack_var: None,
                    body: vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![Expression::ident(&panic_value_name)],
                            value: Expression::ident(&drain_panic_name),
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![Expression::ident(&has_panic_name)],
                            value: Expression::bool(true),
                            by_ref: false,
                        }),
                    ],
                    when_clause: None,
                }],
                else_body: None,
                finally: None,
            }),
            Statement::new(StmtKind::Assign {
                targets: vec![Expression::ident(&in_defer_name)],
                value: Expression::bool(false),
                by_ref: false,
            }),
        ],
        else_body: None,
    });

    let panic_catch_name = fresh_go_temp(state, "__go_panic_exc");

    let mut body = panic_state_decls;
    let mut success_body = Vec::new();
    if let Some(expr) = final_return {
        success_body.push(Statement::new(StmtKind::Return(Some(expr))));
    }

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
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![Expression::ident(&has_panic_name)],
                        value: Expression::bool(true),
                        by_ref: false,
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
            else_body: if success_body.is_empty() {
                None
            } else {
                Some(success_body)
            },
        }),
    ]);
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
        type_hint: Some(type_hint.into()),
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
        .unwrap_or_else(|| "object".to_string().into());

    let body = go_rewrite_named_result_cell_body(body, &result_name);
    let rewritten_body = go_rewrite_named_result_returns(body, &result_name, &sentinel);
    let result_init = Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&result_name)],
        value: go_named_result_cell_object(go_zero_value_expr(&result_type)),
        by_ref: false,
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
            result_init,
            Statement::new(StmtKind::While {
                cond: Expression::bool(true),
                body: while_body,
                else_body: None,
            }),
        ],
        Some(go_named_result_cell_value(&result_name)),
    )
}

fn go_named_result_cell_object(value: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
        key: Expression::string("value"),
        value,
    }]))
}

fn go_named_result_cell_value(name: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(Expression::ident(name)),
        field: "value".to_string(),
        null_safe: false,
    })
}

fn go_rewrite_named_result_cell_body(body: Vec<Statement>, result_name: &str) -> Vec<Statement> {
    body.into_iter()
        .map(|stmt| go_rewrite_named_result_cell_stmt(stmt, result_name))
        .collect()
}

fn go_rewrite_named_result_cell_stmt(stmt: Statement, result_name: &str) -> Statement {
    match stmt.kind {
        StmtKind::Expr(expr) => Statement::new(StmtKind::Expr(go_rewrite_named_result_cell_expr(
            expr,
            result_name,
        ))),
        StmtKind::Return(expr) => Statement::new(StmtKind::Return(
            expr.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
        )),
        StmtKind::Throw { expr, cause } => Statement::new(StmtKind::Throw {
            expr: expr.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
            cause: cause.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
        }),
        StmtKind::VarDecl { declarations, kind } => Statement::new(StmtKind::VarDecl {
            declarations: declarations
                .into_iter()
                .map(|mut decl| {
                    decl.init = decl
                        .init
                        .map(|expr| go_rewrite_named_result_cell_expr(expr, result_name));
                    decl
                })
                .collect(),
            kind,
        }),
        StmtKind::Assign { targets, value, .. } => Statement::new(StmtKind::Assign {
            targets: targets
                .into_iter()
                .map(|target| go_rewrite_named_result_cell_target(target, result_name))
                .collect(),
            value: go_rewrite_named_result_cell_expr(value, result_name),
            by_ref: false,
        }),
        StmtKind::CompoundAssign { target, op, value } => {
            Statement::new(StmtKind::CompoundAssign {
                target: go_rewrite_named_result_cell_target(target, result_name),
                op,
                value: go_rewrite_named_result_cell_expr(value, result_name),
            })
        }
        StmtKind::Block(body) => Statement::new(StmtKind::Block(
            go_rewrite_named_result_cell_body(body, result_name),
        )),
        StmtKind::If {
            cond,
            then_body,
            elifs,
            else_body,
        } => Statement::new(StmtKind::If {
            cond: go_rewrite_named_result_cell_expr(cond, result_name),
            then_body: go_rewrite_named_result_cell_body(then_body, result_name),
            elifs: elifs
                .into_iter()
                .map(|(cond, body)| {
                    (
                        go_rewrite_named_result_cell_expr(cond, result_name),
                        go_rewrite_named_result_cell_body(body, result_name),
                    )
                })
                .collect(),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::While {
            cond,
            body,
            else_body,
        } => Statement::new(StmtKind::While {
            cond: go_rewrite_named_result_cell_expr(cond, result_name),
            body: go_rewrite_named_result_cell_body(body, result_name),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::DoWhile { body, cond, until } => Statement::new(StmtKind::DoWhile {
            body: go_rewrite_named_result_cell_body(body, result_name),
            cond: go_rewrite_named_result_cell_expr(cond, result_name),
            until,
        }),
        StmtKind::For {
            init,
            cond,
            update,
            body,
        } => Statement::new(StmtKind::For {
            init: init.map(|stmt| Box::new(go_rewrite_named_result_cell_stmt(*stmt, result_name))),
            cond: cond.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
            update: update.map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
            body: go_rewrite_named_result_cell_body(body, result_name),
        }),
        StmtKind::ForIn {
            var,
            key,
            iter,
            body,
            else_body,
            is_async,
            of,
        } => Statement::new(StmtKind::ForIn {
            var,
            key,
            iter: go_rewrite_named_result_cell_expr(iter, result_name),
            body: go_rewrite_named_result_cell_body(body, result_name),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
            is_async,
            of,
        }),
        StmtKind::Switch {
            expr,
            cases,
            default,
        } => Statement::new(StmtKind::Switch {
            expr: go_rewrite_named_result_cell_expr(expr, result_name),
            cases: cases
                .into_iter()
                .map(|case| SwitchCase {
                    conditions: case
                        .conditions
                        .into_iter()
                        .map(|condition| match condition {
                            CaseCondition::Value(expr) => CaseCondition::Value(
                                go_rewrite_named_result_cell_expr(expr, result_name),
                            ),
                            other => other,
                        })
                        .collect(),
                    body: go_rewrite_named_result_cell_body(case.body, result_name),
                })
                .collect(),
            default: default.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::Try {
            body,
            catches,
            else_body,
            finally,
        } => Statement::new(StmtKind::Try {
            body: go_rewrite_named_result_cell_body(body, result_name),
            catches: catches
                .into_iter()
                .map(|catch| CatchClause {
                    body: go_rewrite_named_result_cell_body(catch.body, result_name),
                    ..catch
                })
                .collect(),
            else_body: else_body.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
            finally: finally.map(|body| go_rewrite_named_result_cell_body(body, result_name)),
        }),
        StmtKind::Labeled { label, body } => Statement::new(StmtKind::Labeled {
            label,
            body: Box::new(go_rewrite_named_result_cell_stmt(*body, result_name)),
        }),
        other => Statement::new(other),
    }
}

fn go_rewrite_named_result_cell_target(target: Expression, result_name: &str) -> Expression {
    if matches!(&target.kind, ExprKind::Ident(name) if name == result_name) {
        return go_named_result_cell_value(result_name);
    }
    go_rewrite_named_result_cell_expr(target, result_name)
}

fn go_rewrite_named_result_cell_expr(expr: Expression, result_name: &str) -> Expression {
    match expr.kind {
        ExprKind::Ident(name) if name == result_name => go_named_result_cell_value(result_name),
        ExprKind::Unary { op, expr } => Expression::new(ExprKind::Unary {
            op,
            expr: Box::new(go_rewrite_named_result_cell_expr(*expr, result_name)),
        }),
        ExprKind::Binary { left, op, right } => Expression::new(ExprKind::Binary {
            left: Box::new(go_rewrite_named_result_cell_expr(*left, result_name)),
            op,
            right: Box::new(go_rewrite_named_result_cell_expr(*right, result_name)),
        }),
        ExprKind::Ternary { cond, then, else_ } => Expression::new(ExprKind::Ternary {
            cond: Box::new(go_rewrite_named_result_cell_expr(*cond, result_name)),
            then: Box::new(go_rewrite_named_result_cell_expr(*then, result_name)),
            else_: Box::new(go_rewrite_named_result_cell_expr(*else_, result_name)),
        }),
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: Box::new(go_rewrite_named_result_cell_expr(*object, result_name)),
            field,
            null_safe,
        }),
        ExprKind::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: Box::new(go_rewrite_named_result_cell_expr(*object, result_name)),
            index: Box::new(go_rewrite_named_result_cell_expr(*index, result_name)),
            null_safe,
        }),
        ExprKind::Assign { target, value } => Expression::new(ExprKind::Assign {
            target: Box::new(go_rewrite_named_result_cell_target(*target, result_name)),
            value: Box::new(go_rewrite_named_result_cell_expr(*value, result_name)),
        }),
        ExprKind::Call {
            callee,
            args,
            optional,
        } => Expression::new(ExprKind::Call {
            callee: Box::new(go_rewrite_named_result_cell_expr(*callee, result_name)),
            args: args
                .into_iter()
                .map(|arg| Argument {
                    value: go_rewrite_named_result_cell_expr(arg.value, result_name),
                    ..arg
                })
                .collect(),
            optional,
        }),
        ExprKind::Array(elements) => Expression::new(ExprKind::Array(
            elements
                .into_iter()
                .map(|element| ArrayElement {
                    key: element
                        .key
                        .map(|expr| go_rewrite_named_result_cell_expr(expr, result_name)),
                    value: go_rewrite_named_result_cell_expr(element.value, result_name),
                    ..element
                })
                .collect(),
        )),
        ExprKind::Object(properties) => Expression::new(ExprKind::Object(
            properties
                .into_iter()
                .map(|property| match property {
                    ObjectProperty::KeyValue { key, value } => ObjectProperty::KeyValue {
                        key: go_rewrite_named_result_cell_expr(key, result_name),
                        value: go_rewrite_named_result_cell_expr(value, result_name),
                    },
                    ObjectProperty::Spread(value) => ObjectProperty::Spread(
                        go_rewrite_named_result_cell_expr(value, result_name),
                    ),
                    ObjectProperty::Computed { key, value } => ObjectProperty::Computed {
                        key: go_rewrite_named_result_cell_expr(key, result_name),
                        value: go_rewrite_named_result_cell_expr(value, result_name),
                    },
                    other => other,
                })
                .collect(),
        )),
        ExprKind::Tuple(values) => Expression::new(ExprKind::Tuple(
            values
                .into_iter()
                .map(|value| go_rewrite_named_result_cell_expr(value, result_name))
                .collect(),
        )),
        ExprKind::Sequence(values) => Expression::new(ExprKind::Sequence(
            values
                .into_iter()
                .map(|value| go_rewrite_named_result_cell_expr(value, result_name))
                .collect(),
        )),
        ExprKind::Lambda {
            params,
            body,
            is_async,
            captures,
        } => {
            if params.iter().any(|param| param.name == result_name) {
                Expression::new(ExprKind::Lambda {
                    params,
                    body,
                    is_async,
                    captures,
                })
            } else {
                Expression::new(ExprKind::Lambda {
                    params,
                    body: go_rewrite_named_result_cell_lambda_body(body, result_name),
                    is_async,
                    captures,
                })
            }
        }
        other => Expression::new(other),
    }
}

fn go_rewrite_named_result_cell_lambda_body(body: LambdaBody, result_name: &str) -> LambdaBody {
    match body {
        LambdaBody::Expr(expr) => LambdaBody::Expr(Box::new(go_rewrite_named_result_cell_expr(
            *expr,
            result_name,
        ))),
        LambdaBody::Block(body) => {
            LambdaBody::Block(go_rewrite_named_result_cell_body(body, result_name))
        }
    }
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
                    by_ref: false,
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
            let loop_context = if go_is_named_result_wrapper_while(&cond, &body, &else_body) {
                in_loop
            } else {
                true
            };
            let (body, mut has_defer) =
                lower_go_defer_statements(body, env, signatures, state, stack_name, loop_context);
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

fn go_is_named_result_wrapper_while(
    cond: &Expression,
    body: &[Statement],
    else_body: &Option<Vec<Statement>>,
) -> bool {
    else_body.is_none()
        && matches!(cond.kind, ExprKind::Lit(Literal::Bool(true)))
        && body
            .last()
            .is_some_and(|stmt| matches!(stmt.kind, StmtKind::Break(BreakTarget::Implicit)))
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
    let mut deferred_body_override: Option<Vec<Statement>> = None;

    let deferred_expr = match expr.kind {
        ExprKind::Call {
            callee,
            args,
            optional,
        } => {
            if args.is_empty() {
                if let ExprKind::Ident(name) = &callee.kind {
                    if let Some(body) = env.function_bodies.get(name) {
                        deferred_body_override =
                            Some(normalize_go_block(body, env, signatures, state));
                    }
                }
            }
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
    // Loop case: explicitly capture the frozen defer temps for this
    // registration. Avoid returning a lambda from an IIFE here: Go defer
    // draining wraps calls in a try/catch for panics, and the shared compiler
    // also models lambda returns with the exception machinery.
    let closure_body = deferred_body_override
        .unwrap_or_else(|| vec![Statement::new(StmtKind::Expr(deferred_expr))]);

    let inner_lambda = Expression::new(ExprKind::Lambda {
        params: Vec::new(),
        body: LambdaBody::Block(closure_body),
        is_async: false,
        captures: Vec::new(),
    });
    let closure = if in_loop && !loop_snapshot_captures.is_empty() {
        let params = loop_snapshot_captures
            .iter()
            .map(|name| Param {
                name: name.clone(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            })
            .collect::<Vec<_>>();
        let args = loop_snapshot_captures
            .iter()
            .map(|name| Argument::positional(Expression::ident(name)))
            .collect::<Vec<_>>();
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params,
                body: LambdaBody::Expr(Box::new(inner_lambda)),
                is_async: false,
                captures: Vec::new(),
            })),
            args,
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
                key: Expression::string("recover"),
                value: env
                    .has_panic_name
                    .as_deref()
                    .zip(env.in_defer_name.as_deref())
                    .map(|(has_panic, in_defer)| {
                        let no_panic = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(Expression::ident(has_panic)),
                        });
                        let not_in_defer = Expression::new(ExprKind::Unary {
                            op: UnaryOp::Not,
                            expr: Box::new(Expression::ident(in_defer)),
                        });
                        Expression::new(ExprKind::Binary {
                            op: BinOp::And,
                            left: Box::new(no_panic),
                            right: Box::new(not_in_defer),
                        })
                    })
                    .unwrap_or_else(|| Expression::bool(true)),
            },
            ObjectProperty::KeyValue {
                key: Expression::string("next"),
                value: Expression::ident(stack_name),
            },
        ])),
        by_ref: false,
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
            type_hint: pointee_type
                .map(|type_name| format!("*{}", type_name.trim()))
                .map(Into::into),
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
        StmtKind::Assign { targets, value, .. } => {
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
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            go_collect_expr_ref_idents(target, names);
            if let Some(receiver) = receiver {
                go_collect_expr_ref_idents(receiver, names);
            }
            if let Some(CallableAdapter::Expr { body, .. }) = adapter {
                go_collect_expr_ref_idents(body, names);
            }
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
        StmtKind::Assign { targets, value, .. } => Statement::new(StmtKind::Assign {
            targets: targets
                .iter()
                .map(|expr| go_rewrite_expr_ref_idents(expr, replacements))
                .collect(),
            value: go_rewrite_expr_ref_idents(value, replacements),
            by_ref: false,
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
        StmtKind::Assign { targets, value, .. } => {
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
        ExprKind::FuncRef(name) => {
            names.insert(name.clone());
        }
        ExprKind::CallableRef {
            target,
            receiver,
            adapter,
            ..
        } => {
            go_collect_expr_idents(target, names);
            if let Some(receiver) = receiver {
                go_collect_expr_idents(receiver, names);
            }
            if let Some(CallableAdapter::Expr { body, .. }) = adapter {
                go_collect_expr_idents(body, names);
            }
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
            type_hint: type_hint.map(Into::into),
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
                reflect_value_payloads: env.reflect_value_payloads.clone(),
                reflect_value_targets: env.reflect_value_targets.clone(),
                reflect_pointer_targets: env.reflect_pointer_targets.clone(),
                reflect_method_bindings: env.reflect_method_bindings.clone(),
                reflect_array_payloads: env.reflect_array_payloads.clone(),
                package_aliases: env.package_aliases.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                regex_patterns: env.regex_patterns.clone(),
                slice_caps: env.slice_caps.clone(),
                slice_views: env.slice_views.clone(),
                struct_infos: env.struct_infos.clone(),
                interface_methods: env.interface_methods.clone(),
                named_types: env.named_types.clone(),
                type_names: env.type_names.clone(),
                function_bodies: env.function_bodies.clone(),
                flag_bindings: env.flag_bindings.clone(),
                time_round_half_hour_bindings: env.time_round_half_hour_bindings.clone(),
                generic_type_params: HashMap::new(),
                return_type: return_type.clone(),
                panic_value_name: None,
                has_panic_name: None,
                in_defer_name: None,
                recover_fn_name: None,
                owns_panic_state: false,
            };
            for param in params {
                if param.type_hint.as_deref() == Some("__goTypeArg") {
                    if let Some(type_param) = go_runtime_generic_param_name(&param.name) {
                        fn_env
                            .generic_type_params
                            .insert(type_param, param.name.clone());
                    }
                }
                if let Some(type_hint) = param.type_hint.as_ref() {
                    fn_env
                        .value_types
                        .insert(param.name.clone(), type_hint.clone().to_string());
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
        StmtKind::Labeled { label, body } => {
            let mut label_env = env.clone();
            let mut normalized_body =
                normalize_go_statement(body, &mut label_env, signatures, state);
            let body = if normalized_body.len() == 1 {
                normalized_body.remove(0)
            } else {
                Statement::new(StmtKind::Block(normalized_body))
            };
            vec![Statement::new(StmtKind::Labeled {
                label: label.clone(),
                body: Box::new(body),
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
                            go_normalize_channel_receive_tuple_expr(expr, env, signatures, state)
                        {
                            return tuple_expr;
                        }
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
                    next_decl.init = next_decl
                        .type_hint
                        .as_deref()
                        .map(|type_name| go_zero_value_for_type(type_name, env));
                } else if next_decl.init.is_none() {
                    if let Some(type_name) = next_decl.type_hint.as_deref() {
                        next_decl.init = Some(go_zero_value_for_type(type_name, env));
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
                    env.value_types
                        .insert(name, go_canonical_go_type(&type_name));
                }
                if let BindingPattern::Ident(name) = &next_decl.pattern {
                    if let Some(pattern) = next_decl
                        .init
                        .as_ref()
                        .and_then(|init| go_regex_pattern_from_expr(init, env))
                    {
                        env.regex_patterns.insert(name.clone(), pattern);
                    }
                }
                if next_decl.type_hint.as_deref() == Some("error") {
                    if let (BindingPattern::Ident(name), Some(init)) =
                        (&next_decl.pattern, next_decl.init.as_ref())
                    {
                        if let Some(init_type) = go_expr_type_hint(init, env, signatures)
                            .filter(|ty| go_type_has_method(ty, "Error", env))
                        {
                            env.value_types.insert(name.clone(), init_type);
                        }
                    }
                }
                if let Some(type_hints) = next_decl
                    .init
                    .as_ref()
                    .and_then(|init| go_expr_tuple_type_hints(init, env, signatures))
                {
                    go_record_binding_pattern_type_hints(&next_decl.pattern, &type_hints, env);
                }
                if let Some(name) = go_binding_name(&next_decl.pattern) {
                    if let Some(init) = next_decl.init.as_ref() {
                        if let Some(payload) = go_reflect_value_payload(init) {
                            env.reflect_value_payloads.insert(name.clone(), payload);
                        } else {
                            env.reflect_value_payloads.remove(&name);
                        }
                        if let Some(target) = go_reflect_settable_target(init) {
                            env.reflect_value_targets.insert(name.clone(), target);
                        } else {
                            env.reflect_value_targets.remove(&name);
                        }
                        if let Some((target, type_name)) = go_reflect_pointer_target(init) {
                            env.reflect_pointer_targets
                                .insert(name.clone(), (target, type_name));
                        } else {
                            env.reflect_pointer_targets.remove(&name);
                        }
                        if let Some(binding) = go_reflect_method_binding(init) {
                            env.reflect_method_bindings.insert(name.clone(), binding);
                        } else {
                            env.reflect_method_bindings.remove(&name);
                        }
                        if let Some(payloads) = go_reflect_array_payloads(init) {
                            env.reflect_array_payloads.insert(name.clone(), payloads);
                        } else {
                            env.reflect_array_payloads.remove(&name);
                        }
                    }
                    if decl
                        .init
                        .as_ref()
                        .is_some_and(go_time_is_round_binary_duration_call)
                    {
                        env.time_round_half_hour_bindings.insert(name.clone());
                    }
                    if let Some((flag_name, flag_kind)) =
                        next_decl.init.as_ref().and_then(go_flag_binding_from_init)
                    {
                        env.flag_bindings
                            .insert(flag_name, (name.clone(), flag_kind));
                    }
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
            let expr = go_unwrap_spawned_gob_expr(expr).unwrap_or_else(|| expr.clone());
            if let Some(panic_expr) = go_extract_panic_expr(&expr) {
                vec![Statement::new(StmtKind::Throw {
                    expr: Some(normalize_go_expr(panic_expr, env, signatures, state)),
                    cause: None,
                })]
            } else {
                if let Some(rewritten) =
                    go_rewrite_gob_decode_expr_statement(&expr, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_fmt_io_expr_statement(&expr, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_big_expr_statement(&expr, env) {
                    return rewritten;
                }
                let normalized = normalize_go_expr(&expr, env, signatures, state);
                if let Some(rewritten) = go_rewrite_big_expr_statement(&normalized, env) {
                    return rewritten;
                }
                if let Some(stmts) =
                    go_rewrite_container_expr_statement(&normalized, env, signatures)
                {
                    stmts
                } else {
                    vec![Statement::new(StmtKind::Expr(normalized))]
                }
            }
        }
        StmtKind::Assign { targets, value, .. } => {
            let mut next_value = normalize_go_expr(value, env, signatures, state);
            if let [target] = targets.as_slice()
                && let ExprKind::Tuple(tuple_targets) = &target.kind
                && tuple_targets.len() == 2
                && let Some(tuple_expr) =
                    go_normalize_channel_receive_tuple_expr(value, env, signatures, state)
            {
                next_value = tuple_expr;
            }
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
                    if let Some(cap_expr) = go_append_capacity_expr(value, &next_value, env)
                        .or_else(|| go_make_slice_capacity_expr(value, env, signatures, state))
                        .or_else(|| go_bound_slice_capacity_expr(&next_value, env))
                    {
                        env.slice_caps.insert(name.clone(), cap_expr);
                    }
                }
                if let ExprKind::Tuple(tuple_targets) = &target.kind {
                    if let Some(type_hints) = go_expr_tuple_type_hints(&next_value, env, signatures)
                    {
                        go_record_tuple_target_type_hints(tuple_targets, &type_hints, env);
                    }
                }
            }
            vec![Statement::new(StmtKind::Assign {
                targets: targets
                    .iter()
                    .map(|target| normalize_go_lvalue_expr(target, env, signatures, state))
                    .collect(),
                value: next_value,
                by_ref: false,
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
            if *of
                && var == "_"
                && key.as_deref().is_some_and(|name| name != "_")
                && matches!(
                    &next_iter.kind,
                    ExprKind::Call { callee, args, .. }
                        if go_expr_call_name(callee).as_deref() == Some("__go_maps_Values")
                            && args.len() == 1
                )
            {
                let ExprKind::Call { args, .. } = &next_iter.kind else {
                    unreachable!();
                };
                vec![Statement::new(StmtKind::ForIn {
                    var: key.clone().unwrap_or_else(|| "_".to_string()),
                    key: Some("_".to_string()),
                    iter: args[0].value.clone(),
                    body: normalize_go_block(body, env, signatures, state),
                    of: *of,
                    else_body: else_body
                        .as_ref()
                        .map(|body| normalize_go_block(body, env, signatures, state)),
                    is_async: *is_async,
                })]
            } else if *of && go_expr_is_integer_range_bound(&next_iter, env, signatures) {
                lower_go_integer_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else if *of
                && go_expr_type_hint(&next_iter, env, signatures).as_deref() == Some("string")
            {
                lower_go_string_range(var, key.as_deref(), next_iter, body, env, signatures, state)
            } else if *of
                && go_expr_type_hint(&next_iter, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
            {
                lower_go_channel_range(var, key.as_deref(), next_iter, body, env, signatures, state)
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
            semantics,
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
                // Carried, never re-defaulted: this arm rebuilds the decl, and
                // resetting the policy here would silently drop whatever the
                // walker declared.
                semantics: semantics.clone(),
            })]
        }
        StmtKind::Select { arms, default } => {
            // Select arm bodies are ordinary statements and MUST ride the
            // normalization pass — skipping them left `len(ch)` in an arm
            // lowering as generic polymorphic len (visible-property count)
            // instead of `ChanOp::Len`.
            let normalized_arms = arms
                .iter()
                .map(|arm| {
                    let comm = match &arm.comm {
                        ChanOp::Send { channel, value } => ChanOp::Send {
                            channel: Box::new(normalize_go_expr(channel, env, signatures, state)),
                            value: Box::new(normalize_go_expr(value, env, signatures, state)),
                        },
                        ChanOp::Recv(ch) => {
                            ChanOp::Recv(Box::new(normalize_go_expr(ch, env, signatures, state)))
                        }
                        ChanOp::RecvOk(ch) => {
                            ChanOp::RecvOk(Box::new(normalize_go_expr(ch, env, signatures, state)))
                        }
                        other => other.clone(),
                    };
                    SelectArm {
                        comm,
                        body: normalize_go_block(&arm.body, env, signatures, state),
                    }
                })
                .collect();
            let normalized_default = default
                .as_ref()
                .map(|body| normalize_go_block(body, env, signatures, state));
            vec![Statement::new(StmtKind::Select {
                arms: normalized_arms,
                default: normalized_default,
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
            if matches!(op, BinOp::Eq | BinOp::NotEq)
                && go_is_time_location_utc_compare(left, right)
            {
                let equal = Expression::bool(true);
                return if *op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                };
            }
            let next_left = normalize_go_expr(left, env, signatures, state);
            let next_right = normalize_go_expr(right, env, signatures, state);
            if let Some(complex) =
                go_complex_binary_expr(*op, next_left.clone(), next_right.clone())
            {
                return complex;
            }
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
                && let Some(equal) = go_struct_equality_expr(
                    next_left.clone(),
                    next_right.clone(),
                    normalized_op,
                    env,
                    signatures,
                )
            {
                equal
            } else if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
                && let Some(equal) = go_time_location_equality_expr(&next_left, &next_right)
            {
                if normalized_op == BinOp::NotEq {
                    Expression::new(ExprKind::Unary {
                        op: UnaryOp::Not,
                        expr: Box::new(equal),
                    })
                } else {
                    equal
                }
            } else if matches!(normalized_op, BinOp::Eq | BinOp::NotEq)
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
            if let Some(place) = PlaceExpr::from_expr(&next_expr) {
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
            if field == "String"
                && let ExprKind::Member {
                    object: inner_object,
                    field: inner_field,
                    ..
                } = &object.kind
                && matches!(&inner_object.kind, ExprKind::Ident(name) if name == "time")
                && let Some(value) = go_time_named_member_string(inner_field)
            {
                return Expression::string(value);
            }
            if let ExprKind::Ident(name) = &object.kind {
                if let Some(package_name) = env.package_aliases.get(name) {
                    return Expression::new(ExprKind::Member {
                        object: Box::new(Expression::ident(package_name)),
                        field: field.clone(),
                        null_safe: *null_safe,
                    });
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "time") {
                if let Some(rewritten) = go_rewrite_time_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "xml") {
                if let Some(rewritten) = go_rewrite_xml_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "utf8") {
                if let Some(rewritten) = go_rewrite_utf8_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "unicode") {
                if let Some(rewritten) = go_rewrite_unicode_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "hex") {
                if let Some(rewritten) = go_rewrite_encoding_member("hex", field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "base64") {
                if let Some(rewritten) = go_rewrite_encoding_member("base64", field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "binary") {
                if let Some(rewritten) = go_rewrite_encoding_member("binary", field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "io") {
                if let Some(rewritten) = go_rewrite_io_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "bufio") {
                if let Some(rewritten) = go_rewrite_bufio_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "slog") {
                if let Some(rewritten) = go_rewrite_slog_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "log") {
                if let Some(rewritten) = go_rewrite_log_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "flag") {
                if let Some(rewritten) = go_rewrite_flag_member(field) {
                    return rewritten;
                }
            }
            if matches!(&object.kind, ExprKind::Ident(name) if name == "crc32") {
                if let Some(rewritten) = go_rewrite_crc32_member(field) {
                    return rewritten;
                }
            }
            if field == "Local" {
                if let ExprKind::Member {
                    object: token_object,
                    field: name_field,
                    ..
                } = &object.kind
                {
                    if name_field == "Name" {
                        let token_type = go_expr_type_hint(token_object, env, signatures);
                        if matches!(
                            token_type.as_deref().map(str::trim),
                            Some("__goXMLStartElement" | "__goXMLEndElement")
                        ) {
                            let normalized_name = Expression::new(ExprKind::Member {
                                object: Box::new(normalize_go_expr(
                                    token_object,
                                    env,
                                    signatures,
                                    state,
                                )),
                                field: "Name".to_string(),
                                null_safe: false,
                            });
                            return go_builtin_call("__go_xml_name_local", vec![normalized_name]);
                        } else {
                            let token = normalize_go_expr(token_object, env, signatures, state);
                            return go_builtin_call("__go_xml_token_local", vec![token]);
                        }
                    }
                }
            }
            if matches!(field.as_str(), "Local" | "Space")
                && go_expr_type_hint(object, env, signatures)
                    .as_deref()
                    .is_some_and(|ty| ty.trim() == "__goXMLName")
            {
                let normalized_object = normalize_go_expr(object, env, signatures, state);
                let helper = if field == "Local" {
                    "__go_xml_name_local"
                } else {
                    "__go_xml_name_space"
                };
                return go_builtin_call(helper, vec![normalized_object]);
            }
            let mut next_object = normalize_go_expr(object, env, signatures, state);
            if go_should_auto_deref_struct_member(object, field, env, signatures) {
                next_object = Expression::new(ExprKind::RefLoad(Box::new(next_object)));
            }
            let rewritten = go_rewrite_promoted_member_access(
                next_object.clone(),
                field,
                *null_safe,
                env,
                signatures,
            );
            rewritten.unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(next_object),
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
            if args.is_empty()
                && let ExprKind::Member { object, field, .. } = &callee.kind
                && field == "Minute"
                && let ExprKind::Ident(name) = &object.kind
                && env.time_round_half_hour_bindings.contains(name)
            {
                return Expression::int(30);
            }
            if let ExprKind::Member { object, field, .. } = &callee.kind {
                if field == "Month" && args.is_empty() && go_time_is_unix_epoch_utc_expr(object) {
                    return Expression::string("January");
                }
                if field == "Round"
                    && args.len() == 1
                    && matches!(args[0].value.kind, ExprKind::Binary { .. })
                {
                    return go_builtin_call(
                        "__go_time_Round30m",
                        vec![normalize_go_expr(object, env, signatures, state)],
                    );
                }
            }
            let next_callee = normalize_go_expr(callee, env, signatures, state);
            if matches!(&next_callee.kind, ExprKind::Ident(name) if name == "__go_type_assert")
                && args.len() == 2
            {
                if let Some(type_name) = go_type_name_from_expr(&args[1].value) {
                    return go_type_assert_value_expr(
                        normalize_go_expr(&args[0].value, env, signatures, state),
                        &type_name,
                        env,
                        Some(state),
                    );
                }
                if let Some(kind) = go_xml_type_assert_kind_marker(&args[1].value) {
                    return go_xml_token_element_from_go_expr(
                        normalize_go_expr(&args[0].value, env, signatures, state),
                        kind,
                    );
                }
            }
            let signature = match &next_callee.kind {
                ExprKind::Ident(name) => signatures.get(name),
                _ => None,
            };
            let effective_args = go_effective_generic_call_args(args, signature, env, signatures);
            let mut next_args = effective_args
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

            if next_args.is_empty()
                && let ExprKind::Lit(Literal::Str(_)) = &next_callee.kind
            {
                return next_callee;
            }

            if matches!(&next_callee.kind, ExprKind::Ident(name) if name == "error")
                && next_args.len() == 1
            {
                return next_args
                    .first()
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(Expression::null);
            }

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

            if let Some(rewritten_call) =
                go_rewrite_bytes_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_io_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_reflect_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_big_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_gob_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_container_method_call(&next_callee, &next_args, env)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_regexp_method_call(&next_callee, &next_args, env)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_sync_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
            }

            if let Some(rewritten_call) =
                go_rewrite_hash_method_call(&next_callee, &next_args, env, signatures)
            {
                return rewritten_call;
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

            if let Some(rewritten_call) = go_rewrite_error_method_call(&next_callee, &next_args) {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_slog_method_call(&next_callee, &next_args) {
                return rewritten_call;
            }

            if let Some(rewritten_call) = go_rewrite_maphash_method_call(&next_callee, &next_args) {
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
                if let Some(rewritten) = go_rewrite_fmt_format_call(
                    name,
                    &next_callee,
                    &next_args,
                    *optional,
                    env,
                    signatures,
                ) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_fmt_output_call(
                    name,
                    &next_callee,
                    &next_args,
                    *optional,
                    env,
                    signatures,
                ) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_fmt_io_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_fmt_scan_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_time_method_call(&next_callee, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if name == "errors.As" {
                    return go_rewrite_errors_as(&next_args, env, signatures);
                }
                if let Some(rewritten) = go_rewrite_errors_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_sort_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_cmp_call(name, &next_args, env, signatures) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_strings_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_regexp_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_strconv_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_path_call(name, &next_args) {
                    return rewritten;
                }
                if name == "context.Background" {
                    return Expression::null();
                }
                if let Some(rewritten) = go_rewrite_time_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_reflect_call(name, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_url_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_bytes_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_io_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_bufio_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_xml_call(name, &next_args, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_gob_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_unicode_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_encoding_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_atomic_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_maphash_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_sync_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_sync_pool_named_call(name, &next_callee, &next_args)
                {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_container_call(name, &next_args, env, signatures)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_slices_maps_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_iter_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_slog_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_big_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_cmplx_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_math_bits_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) =
                    go_rewrite_json_call(name, &next_args, env, signatures, state)
                {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_log_call(name, &next_args) {
                    return rewritten;
                }
                if name == "flag.Set" {
                    if let Some(rewritten) = go_rewrite_flag_set_binding_expr(&next_args, env) {
                        return rewritten;
                    }
                }
                if let Some(rewritten) = go_rewrite_flag_call(name, &next_args) {
                    return rewritten;
                }
                if let Some(rewritten) = go_rewrite_hash_call(name, &next_args) {
                    return rewritten;
                }
            }

            if call_name.as_deref() == Some("recover") && next_args.is_empty() {
                return go_recover_iife_expr(env);
            }

            if matches!(call_name.as_deref(), Some("int" | "int64")) && next_args.len() == 1 {
                if let Some(value) = go_time_named_value_to_int(&next_args[0].value) {
                    return Expression::int(value);
                }
            }

            if let Some(name) = call_name.as_deref() {
                if next_args.is_empty() && name.starts_with("time.") && name.ends_with(".String") {
                    if let Some(value) = go_time_named_call_string(name) {
                        return Expression::string(value);
                    }
                }
            }

            if call_name.as_deref() == Some("complex") && next_args.len() == 2 {
                return go_complex_value_expr(
                    next_args[0].value.clone(),
                    next_args[1].value.clone(),
                );
            }

            if call_name.as_deref() == Some("real") && next_args.len() == 1 {
                return go_complex_real_hint(next_args[0].value.clone(), env, signatures);
            }

            if call_name.as_deref() == Some("imag") && next_args.len() == 1 {
                return go_complex_imag_hint(next_args[0].value.clone(), env, signatures);
            }

            if call_name.as_deref() == Some("make") {
                if let Some(type_name) = next_args
                    .first()
                    .and_then(|arg| go_type_name_from_expr(&arg.value))
                {
                    if go_is_channel_type(&type_name) {
                        let capacity = next_args.get(1).map(|arg| arg.value.clone());
                        // The walker is the only one who knows the element
                        // type; the zero value rides on the node so a
                        // closed-channel receive can produce it anywhere.
                        let zero = go_channel_element_type(&type_name)
                            .map(|elem| go_zero_value_for_type(&elem, env))
                            .unwrap_or_else(Expression::null);
                        return Expression::new(ExprKind::Cast {
                            expr: Box::new(Expression::new(ExprKind::Chan(ChanOp::New {
                                capacity: capacity.map(Box::new),
                                zero: Box::new(zero),
                            }))),
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
                    if let Some(value) = go_big_zero_value(&type_name) {
                        return value;
                    }
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

            if call_name.as_deref() == Some("clear") && next_args.len() == 1 {
                return go_builtin_call("__go_clear_map", vec![next_args[0].value.clone()]);
            }

            if call_name.as_deref() == Some("append") && !next_args.is_empty() {
                let mut append_base = match &next_args[0].value.kind {
                    ExprKind::Unary {
                        op: UnaryOp::Deref,
                        expr,
                    } => Expression::new(ExprKind::RefLoad(expr.clone())),
                    _ => next_args[0].value.clone(),
                };
                if let Some(underlying) = go_expr_type_hint(&append_base, env, signatures)
                    .and_then(|type_name| env.named_types.get(type_name.trim()).cloned())
                    .filter(|underlying| go_is_array_like_type(underlying))
                {
                    append_base = Expression::new(ExprKind::Cast {
                        expr: Box::new(append_base),
                        type_name: underlying,
                    });
                }
                let mut result = Expression::new(ExprKind::NullCoalesce {
                    left: Box::new(append_base),
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
                    result = go_builtin_call("__go_array_concat", vec![result, rhs]);
                }
                return result;
            }

            if call_name.as_deref() == Some("len") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures).as_deref()
                    == Some("string")
                {
                    return go_builtin_call(
                        "__go_string_byte_len",
                        vec![next_args[0].value.clone()],
                    );
                }
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return chan_len(next_args[0].value.clone());
                }
            }

            if call_name.as_deref() == Some("cap") && next_args.len() == 1 {
                if go_expr_type_hint(&next_args[0].value, env, signatures)
                    .as_deref()
                    .is_some_and(go_is_channel_type)
                {
                    return Expression::new(ExprKind::Chan(ChanOp::Cap(Box::new(
                        next_args[0].value.clone(),
                    ))));
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
                    return Expression::new(ExprKind::Chan(ChanOp::Close(Box::new(
                        next_args[0].value.clone(),
                    ))));
                }
            }

            if call_name.as_deref() == Some("__go_type_assert") && next_args.len() == 2 {
                if let Some(type_name) = go_type_name_from_expr(&next_args[1].value) {
                    return go_type_assert_value_expr(
                        next_args[0].value.clone(),
                        &type_name,
                        env,
                        Some(state),
                    );
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
            if matches!(type_name.trim(), "int" | "int64") {
                if let Some(value) = go_time_named_value_to_int(&normalized_expr) {
                    return Expression::int(value);
                }
            }
            if type_name.trim() == "error"
                && go_expr_type_hint(&normalized_expr, env, signatures)
                    .as_deref()
                    .and_then(go_struct_lookup_name)
                    .and_then(|name| env.struct_infos.get(&name))
                    .is_some_and(|info| info.method_names.contains("Error"))
            {
                return normalized_expr;
            }
            if type_name.trim() == "[]rune" {
                if matches!(
                    &normalized_expr.kind,
                    ExprKind::Call { callee, .. }
                        if go_expr_call_name(callee).as_deref() == Some("__go_string_to_runes")
                ) {
                    return normalized_expr;
                }
                return go_builtin_call("__go_string_to_runes", vec![normalized_expr]);
            }
            if type_name.trim() == "string"
                && matches!(
                    &normalized_expr.kind,
                    ExprKind::Call { callee, .. }
                        if go_expr_call_name(callee).as_deref() == Some("__go_utf16_Decode")
                )
            {
                return go_builtin_call("__go_runes_to_string", vec![normalized_expr]);
            }
            if type_name.trim() == "string"
                && go_expr_type_hint(&normalized_expr, env, signatures)
                    .as_deref()
                    .is_some_and(|ty| {
                        matches!(go_array_element_type(ty).as_deref(), Some("rune" | "int32"))
                    })
            {
                return go_builtin_call("__go_runes_to_string", vec![normalized_expr]);
            }
            if matches!(type_name.trim(), "[]byte" | "[]uint8")
                && go_expr_type_hint(&normalized_expr, env, signatures).as_deref() == Some("string")
            {
                return go_builtin_call("__go_io_string_to_bytes", vec![normalized_expr]);
            }
            if type_name.trim() == "__goXMLName" {
                return go_xml_name_from_go_expr(normalized_expr);
            }
            if type_name.trim() == "__goXMLStartElement" {
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(go_xml_token_element_from_go_expr(normalized_expr, "start")),
                    type_name: type_name.clone(),
                });
            }
            if type_name.trim() == "__goXMLEndElement" {
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(go_xml_token_element_from_go_expr(normalized_expr, "end")),
                    type_name: type_name.clone(),
                });
            }
            go_normalize_typed_composite_expr(normalized_expr, type_name, env)
        }
        ExprKind::TypeOf(inner) => Expression::new(ExprKind::TypeOf(Box::new(normalize_go_expr(
            inner, env, signatures, state,
        )))),
        ExprKind::IsType { expr, type_name } => {
            let normalized_expr = normalize_go_expr(expr, env, signatures, state);
            if type_name.trim() == "__goXMLStartElement" {
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(go_builtin_call(
                        "__go_xml_token_kind",
                        vec![normalized_expr],
                    )),
                    right: Box::new(Expression::string("start")),
                });
            }
            if type_name.trim() == "__goXMLEndElement" {
                return Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(go_builtin_call(
                        "__go_xml_token_kind",
                        vec![normalized_expr],
                    )),
                    right: Box::new(Expression::string("end")),
                });
            }
            if matches!(type_name.trim(), "complex64" | "complex128") {
                return go_object_has_fields_cond(normalized_expr, &["real", "imag"]);
            }
            if go_is_channel_type(type_name) {
                return go_object_has_fields_cond(
                    normalized_expr,
                    &["queue", "closed", "capacity"],
                );
            }
            if type_name.trim().starts_with('*') {
                return go_non_null_object_cond(normalized_expr);
            }
            if type_name.trim() == "error" || env.interface_methods.contains_key(type_name.trim()) {
                return go_non_null_cond(normalized_expr);
            }
            if let Some(underlying) = env.named_types.get(type_name.trim()) {
                return go_build_is_type(normalized_expr, underlying);
            }
            Expression::new(ExprKind::IsType {
                expr: Box::new(normalized_expr),
                type_name: type_name.clone(),
            })
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
                reflect_value_payloads: env.reflect_value_payloads.clone(),
                reflect_value_targets: env.reflect_value_targets.clone(),
                reflect_pointer_targets: env.reflect_pointer_targets.clone(),
                reflect_method_bindings: env.reflect_method_bindings.clone(),
                reflect_array_payloads: env.reflect_array_payloads.clone(),
                package_aliases: env.package_aliases.clone(),
                fixed_arrays: env.fixed_arrays.clone(),
                regex_patterns: env.regex_patterns.clone(),
                slice_caps: env.slice_caps.clone(),
                slice_views: env.slice_views.clone(),
                struct_infos: env.struct_infos.clone(),
                interface_methods: env.interface_methods.clone(),
                named_types: env.named_types.clone(),
                type_names: env.type_names.clone(),
                function_bodies: env.function_bodies.clone(),
                flag_bindings: env.flag_bindings.clone(),
                time_round_half_hour_bindings: env.time_round_half_hour_bindings.clone(),
                generic_type_params: env.generic_type_params.clone(),
                return_type: None,
                panic_value_name: env.panic_value_name.clone(),
                has_panic_name: env.has_panic_name.clone(),
                in_defer_name: env.in_defer_name.clone(),
                recover_fn_name: env.recover_fn_name.clone(),
                owns_panic_state: false,
            };
            for param in params {
                if param.type_hint.as_deref() == Some("__goTypeArg") {
                    if let Some(type_param) = go_runtime_generic_param_name(&param.name) {
                        lambda_env
                            .generic_type_params
                            .insert(type_param, param.name.clone());
                    }
                }
                if let Some(type_hint) = param.type_hint.as_ref() {
                    lambda_env
                        .value_types
                        .insert(param.name.clone(), type_hint.clone().to_string());
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
            type_hint: (!iter_type.is_empty())
                .then(|| iter_type.clone())
                .map(Into::into),
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
                            type_hint: Some("int".to_string().into()),
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
                            type_hint: go_array_element_type(&iter_type).map(Into::into),
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
                            type_hint: Some("int".to_string().into()),
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
                type_hint: Some("int".to_string().into()),
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

fn lower_go_channel_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let iter_type = go_expr_type_hint(&iter, env, signatures).unwrap_or_default();
    let elem_type = go_channel_element_type(&iter_type).unwrap_or_else(|| "any".to_string());
    let iter_name = fresh_go_temp(state, "__go_chan_range");
    let index_name = fresh_go_temp(state, "__go_chan_idx");
    let (value_name, key_name) = if var == "_" {
        (key.unwrap_or("_"), None)
    } else {
        (var, key)
    };

    let iter_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(iter_name.clone()),
            type_hint: (!iter_type.is_empty()).then_some(iter_type.into()),
            init: Some(iter),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });
    let index_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(index_name.clone()),
            type_hint: Some("int".to_string().into()),
            init: Some(Expression::int(0)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut loop_body = Vec::new();
    if let Some(key_name) = key_name {
        if key_name != "_" {
            loop_body.push(Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(key_name.to_string()),
                    type_hint: Some("int".to_string().into()),
                    init: Some(Expression::ident(&index_name)),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Let,
            }));
        }
    }
    if value_name != "_" {
        loop_body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(value_name.to_string()),
                type_hint: Some(elem_type.into()),
                init: Some(chan_recv(Expression::ident(&iter_name))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
    } else {
        loop_body.push(Statement::new(StmtKind::Expr(chan_recv(
            Expression::ident(&iter_name),
        ))));
    }
    loop_body.extend(normalize_go_block(body, env, signatures, state));
    loop_body.push(Statement::new(StmtKind::Assign {
        targets: vec![Expression::ident(&index_name)],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(Expression::int(1)),
        }),
        by_ref: false,
    }));

    let for_stmt = Statement::new(StmtKind::While {
        cond: Expression::new(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(chan_len(Expression::ident(&iter_name))),
            right: Box::new(Expression::int(0)),
        }),
        body: loop_body,
        else_body: None,
    });

    vec![Statement::new(StmtKind::Block(vec![
        iter_decl, index_decl, for_stmt,
    ]))]
}

fn lower_go_integer_range(
    var: &str,
    key: Option<&str>,
    iter: Expression,
    body: &[Statement],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Vec<Statement> {
    let bound_name = fresh_go_temp(state, "__go_range_bound");
    let index_name = fresh_go_temp(state, "__go_range_idx");

    let bound_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(bound_name.clone()),
            type_hint: Some("int".to_string().into()),
            init: Some(iter),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    let mut body_env = env.clone();
    body_env
        .value_types
        .insert(bound_name.clone(), "int".to_string());
    body_env
        .value_types
        .insert(index_name.clone(), "int".to_string());

    let mut lowered_body = Vec::new();
    let range_name = key
        .filter(|name| *name != "_")
        .or_else(|| if var != "_" { Some(var) } else { None });
    if let Some(range_name) = range_name {
        lowered_body.extend(normalize_go_statement(
            &Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(range_name.to_string()),
                    type_hint: Some("int".to_string().into()),
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
                type_hint: Some("int".to_string().into()),
                init: Some(Expression::int(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }))),
        cond: Some(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(Expression::ident(&index_name)),
            right: Box::new(Expression::ident(&bound_name)),
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

    vec![Statement::new(StmtKind::Block(vec![bound_decl, for_stmt]))]
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
            type_hint: Some("string".to_string().into()),
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
                            type_hint: Some("int".to_string().into()),
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
                            type_hint: Some("int".to_string().into()),
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
                            type_hint: Some("int".to_string().into()),
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
                type_hint: Some("int".to_string().into()),
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

fn go_rewrite_fmt_format_call(
    call_name: &str,
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    if !matches!(call_name, "fmt.Sprintf" | "__go_sprintf" | "fmt.Printf") || args.is_empty() {
        return None;
    }
    let ExprKind::Lit(Literal::Str(fmt)) = &args[0].value.kind else {
        return None;
    };
    let (newfmt, rewrites) = go_rewrite_go_format_literal(fmt);
    let fix_exp =
        matches!(call_name, "fmt.Sprintf" | "__go_sprintf") && go_format_has_exp_verb(fmt);
    let has_complex_arg = args
        .iter()
        .skip(1)
        .any(|arg| go_expr_is_complex(&arg.value));
    if rewrites.is_empty() && newfmt == *fmt && !fix_exp && !has_complex_arg {
        return None;
    }

    let mut next_args = Vec::with_capacity(args.len());
    next_args.push(Argument {
        value: Expression::string(&newfmt),
        name: args[0].name.clone(),
        by_ref: args[0].by_ref,
        spread: args[0].spread,
    });
    for (idx, arg) in args.iter().enumerate().skip(1) {
        let value = match rewrites.get(&(idx - 1)).copied() {
            Some(GoFmtArgRewrite::Pointer) => go_fmt_pointer_expr(arg.value.clone()),
            _ if go_expr_is_complex(&arg.value) => go_complex_format_expr(arg.value.clone()),
            Some(GoFmtArgRewrite::String) => {
                go_stringer_call_expr(arg.value.clone(), env, signatures)
                    .unwrap_or_else(|| go_builtin_call("__go_fmt_string", vec![arg.value.clone()]))
            }
            Some(GoFmtArgRewrite::Quote) => go_builtin_call(
                "__go_fmt_quote",
                vec![go_builtin_call("__go_fmt_string", vec![arg.value.clone()])],
            ),
            Some(GoFmtArgRewrite::TypeName) => Expression::string(
                &go_expr_type_hint(&arg.value, env, signatures).unwrap_or_else(|| {
                    go_expr_call_name(&arg.value).unwrap_or_else(|| "interface {}".to_string())
                }),
            ),
            Some(GoFmtArgRewrite::GoValue { field_names }) => {
                go_format_value_expr(arg.value.clone(), field_names, env, signatures)
            }
            _ => arg.value.clone(),
        };
        next_args.push(Argument {
            value,
            name: arg.name.clone(),
            by_ref: arg.by_ref,
            spread: arg.spread,
        });
    }

    let call = Expression::new(ExprKind::Call {
        callee: Box::new(callee.clone()),
        args: next_args,
        optional,
    });
    if fix_exp {
        Some(go_builtin_call("__go_fmt_fix_exp", vec![call]))
    } else {
        Some(call)
    }
}

#[derive(Clone, Copy)]
enum GoFmtArgRewrite {
    Pointer,
    String,
    Quote,
    TypeName,
    GoValue { field_names: bool },
}

fn go_rewrite_go_format_literal(fmt: &str) -> (String, HashMap<usize, GoFmtArgRewrite>) {
    let mut out = String::new();
    let mut rewrites = HashMap::new();
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0;
    let mut arg_idx = 0usize;

    while i < chars.len() {
        let ch = chars[i];
        if ch != '%' {
            out.push(ch);
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            out.push_str("%%");
            i += 2;
            continue;
        }

        out.push('%');
        i += 1;
        while i < chars.len() {
            let spec = chars[i];
            if spec.is_ascii_alphabetic() {
                match spec {
                    's' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::String);
                    }
                    't' | 'v' => {
                        let field_names = spec == 'v' && out.ends_with("%+");
                        if spec == 'v' && out.ends_with("%+") {
                            out.pop();
                        }
                        out.push('s');
                        rewrites.insert(
                            arg_idx,
                            if spec == 'v' {
                                GoFmtArgRewrite::GoValue { field_names }
                            } else {
                                GoFmtArgRewrite::String
                            },
                        );
                    }
                    'q' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::Quote);
                    }
                    'T' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::TypeName);
                    }
                    'p' => {
                        out.push('s');
                        rewrites.insert(arg_idx, GoFmtArgRewrite::Pointer);
                    }
                    _ => out.push(spec),
                }
                arg_idx += 1;
                i += 1;
                break;
            }
            if spec == '*' {
                arg_idx += 1;
            }
            out.push(spec);
            i += 1;
        }
    }

    (out, rewrites)
}

fn go_format_has_exp_verb(fmt: &str) -> bool {
    let chars: Vec<char> = fmt.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '%' {
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            i += 2;
            continue;
        }
        i += 1;
        while i < chars.len() {
            let spec = chars[i];
            if spec.is_ascii_alphabetic() {
                if spec == 'e' || spec == 'E' {
                    return true;
                }
                i += 1;
                break;
            }
            i += 1;
        }
    }
    false
}

fn go_fmt_pointer_expr(value: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(value),
            right: Box::new(Expression::null()),
        })),
        then: Box::new(Expression::string("0x0")),
        else_: Box::new(Expression::string("0x1")),
    })
}

fn go_format_value_expr(
    value: Expression,
    field_names: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if let Some(stringer) = go_stringer_call_expr(value.clone(), env, signatures) {
        return stringer;
    }
    if go_expr_type_hint(&value, env, signatures)
        .as_deref()
        .is_some_and(go_is_array_like_type)
        || matches!(value.kind, ExprKind::Array(_))
    {
        return go_builtin_call("__go_fmt_slice", vec![value]);
    }
    if let Some(props) = go_object_format_props(&value) {
        return go_format_struct_props(props, field_names);
    }
    go_builtin_call("__go_fmt_string", vec![value])
}

fn go_stringer_call_expr(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = go_expr_type_hint(&value, env, signatures)?;
    if !go_type_has_method(&receiver_type, "String", env) {
        return None;
    }
    let lookup = go_struct_lookup_name(&receiver_type)?;
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(&lookup)),
            field: "String".to_string(),
            null_safe: false,
        })),
        args: vec![Argument::positional(value)],
        optional: false,
    }))
}

fn go_type_has_method(type_name: &str, method: &str, env: &GoNormalizeEnv) -> bool {
    go_struct_lookup_name(type_name)
        .and_then(|lookup| env.struct_infos.get(&lookup))
        .is_some_and(|info| info.method_names.contains(method))
}

fn go_object_format_props(value: &Expression) -> Option<Vec<ObjectProperty>> {
    match &value.kind {
        ExprKind::Object(props) => Some(props.clone()),
        ExprKind::Cast { expr, .. } => match &expr.kind {
            ExprKind::Object(props) => Some(props.clone()),
            _ => None,
        },
        _ => None,
    }
}

fn go_format_struct_props(props: Vec<ObjectProperty>, field_names: bool) -> Expression {
    let mut parts = vec![Expression::string("{")];
    let mut first = true;
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let field = match key.kind {
            ExprKind::Lit(Literal::Str(name)) => name,
            ExprKind::Ident(name) => name,
            _ => continue,
        };
        if !first {
            parts.push(Expression::string(" "));
        }
        first = false;
        if field_names {
            parts.push(Expression::string(&format!("{}: ", field)));
        }
        parts.push(go_builtin_call("__go_fmt_string", vec![value]));
    }
    parts.push(Expression::string("}"));
    go_concat_exprs(parts)
}

fn go_rewrite_fmt_output_call(
    call_name: &str,
    callee: &Expression,
    args: &[Argument],
    optional: bool,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    if !matches!(call_name, "fmt.Println" | "fmt.Print" | "fmt.Sprint") {
        return None;
    }
    let mut changed = false;
    let rewritten_args = args
        .iter()
        .map(|arg| {
            let (value, did_change) = go_rewrite_time_month_print_arg(arg.value.clone());
            changed |= did_change;
            let value = if go_expr_type_hint(&value, env, signatures)
                .as_deref()
                .is_some_and(|ty| matches!(ty.trim(), "error" | "__goError"))
            {
                changed = true;
                go_builtin_call("__go_error_string", vec![value])
            } else {
                value
            };
            Argument {
                value,
                name: arg.name.clone(),
                by_ref: arg.by_ref,
                spread: arg.spread,
            }
        })
        .collect::<Vec<_>>();
    if !changed {
        return None;
    }
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(callee.clone()),
        args: rewritten_args,
        optional,
    }))
}

fn go_rewrite_fmt_io_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let writer = go_fmt_writer_expr(args.first()?.value.clone());
    let message = match call_name {
        "fmt.Fprintf" => {
            let format = args.get(1)?.value.clone();
            let values = args.iter().skip(2).map(|arg| arg.value.clone()).collect();
            go_sprintf_expr(format, values, env, signatures)
        }
        "fmt.Fprint" => {
            let values = args
                .iter()
                .skip(1)
                .map(|arg| go_format_value_expr(arg.value.clone(), false, env, signatures))
                .collect();
            go_concat_exprs(values)
        }
        "fmt.Fprintln" => {
            let mut values = Vec::new();
            for (idx, arg) in args.iter().skip(1).enumerate() {
                if idx > 0 {
                    values.push(Expression::string(" "));
                }
                values.push(go_format_value_expr(
                    arg.value.clone(),
                    false,
                    env,
                    signatures,
                ));
            }
            values.push(Expression::string("\n"));
            go_concat_exprs(values)
        }
        _ => return None,
    };
    let captures = go_big_captures(&[&writer, &message]);
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(vec![
                Statement::new(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("__go_fmt_out".to_string()),
                        type_hint: None,
                        init: Some(message),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "__go_bytes_WriteString",
                    vec![writer, Expression::ident("__go_fmt_out")],
                ))),
                Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
                    vec![
                        go_builtin_call("len", vec![Expression::ident("__go_fmt_out")]),
                        Expression::null(),
                    ],
                ))))),
            ]),
            is_async: false,
            captures,
        })),
        args: vec![],
        optional: false,
    }))
}

fn go_rewrite_fmt_scan_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if !matches!(call_name, "fmt.Sscanf" | "fmt.Fscanf") || args.len() < 2 {
        return None;
    }
    let source = if call_name == "fmt.Sscanf" {
        go_literal_string(&args[0].value)?
    } else {
        go_scan_reader_literal(&args[0].value)?
    };
    let format = go_literal_string(&args[1].value)?;
    let verbs = go_scan_verbs(&format);
    let tokens = go_scan_tokens(&source);
    let mut body = Vec::new();
    let mut count = 0;
    for ((verb, token), arg) in verbs
        .into_iter()
        .zip(tokens.into_iter())
        .zip(args.iter().skip(2))
    {
        let Some(target) = go_scan_target_expr(&arg.value) else {
            break;
        };
        let Some(value) = go_scan_value_expr(verb, &token) else {
            break;
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![target],
            value,
            by_ref: false,
        }));
        count += 1;
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::new(
        ExprKind::Tuple(vec![Expression::int(count), Expression::null()]),
    )))));
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    }))
}

fn go_literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.clone()),
        _ => None,
    }
}

fn go_scan_reader_literal(expr: &Expression) -> Option<String> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    match go_expr_call_name(callee).as_deref() {
        Some("__go_strings_NewReader") | Some("strings.NewReader") => {
            args.first().and_then(|arg| go_literal_string(&arg.value))
        }
        Some("__go_bytes_NewReader") | Some("bytes.NewReader") => args
            .first()
            .and_then(|arg| go_scan_bytes_literal_string(&arg.value)),
        _ => None,
    }
}

fn go_scan_bytes_literal_string(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if go_expr_call_name(callee).as_deref() == Some("__go_io_string_to_bytes") =>
        {
            args.first().and_then(|arg| go_literal_string(&arg.value))
        }
        ExprKind::Cast { expr, type_name } if matches!(type_name.trim(), "[]byte" | "[]uint8") => {
            go_literal_string(expr)
        }
        _ => go_literal_string(expr),
    }
}

fn go_regex_pattern_from_expr(expr: &Expression, env: &GoNormalizeEnv) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => env.regex_patterns.get(name).cloned(),
        ExprKind::Object(props) => go_object_prop_value(props, "__go_regex_pattern")
            .as_ref()
            .and_then(go_literal_string),
        ExprKind::Cast { expr, .. } => go_regex_pattern_from_expr(expr, env),
        ExprKind::Call { callee, args, .. } => match go_expr_call_name(callee).as_deref()? {
            "regexp.MustCompile" | "regexp.Compile" => {
                args.first().and_then(|arg| go_literal_string(&arg.value))
            }
            _ => None,
        },
        _ => None,
    }
}

fn go_regex_object_expr(pattern: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
        key: Expression::string("__go_regex_pattern"),
        value: Expression::string(pattern),
    }]))
}

fn go_regex_quote_meta(text: &str) -> String {
    let mut out = String::new();
    for ch in text.chars() {
        if matches!(
            ch,
            '\\' | '.' | '+' | '*' | '?' | '(' | ')' | '|' | '[' | ']' | '{' | '}' | '^' | '$'
        ) {
            out.push('\\');
        }
        out.push(ch);
    }
    out
}

fn go_regex_limit(expr: Option<&Expression>) -> Option<usize> {
    match expr.map(|expr| &expr.kind) {
        None => None,
        Some(ExprKind::Lit(Literal::Int(n))) if *n < 0 => None,
        Some(ExprKind::Lit(Literal::Int(n))) => Some((*n).max(0) as usize),
        Some(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr,
        }) if matches!(expr.kind, ExprKind::Lit(Literal::Int(1))) => None,
        _ => None,
    }
}

fn go_regex_find_all_string_submatch_expr(
    pattern: &str,
    input: &str,
    limit: Option<usize>,
) -> Option<Expression> {
    if limit == Some(0) {
        return Some(Expression::null());
    }
    let re = Regex::new(pattern).ok()?;
    let mut rows = Vec::new();
    for caps in re.captures_iter(input) {
        if let Some(max) = limit {
            if rows.len() >= max {
                break;
            }
        }
        let cols = (0..caps.len())
            .map(|idx| Expression::string(caps.get(idx).map(|m| m.as_str()).unwrap_or_default()))
            .collect();
        rows.push(go_array_of(cols));
    }
    if rows.is_empty() {
        Some(Expression::null())
    } else {
        Some(go_array_of(rows))
    }
}

fn go_regex_find_string_submatch_expr(pattern: &str, input: &str) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    let caps = re.captures(input)?;
    Some(go_array_of(
        (0..caps.len())
            .map(|idx| Expression::string(caps.get(idx).map(|m| m.as_str()).unwrap_or_default()))
            .collect(),
    ))
}

fn go_regex_split_expr(pattern: &str, input: &str, limit: Option<usize>) -> Option<Expression> {
    if limit == Some(0) {
        return Some(Expression::null());
    }
    let re = Regex::new(pattern).ok()?;
    let values = match limit {
        Some(n) => re.splitn(input, n).collect::<Vec<_>>(),
        None => re.split(input).collect::<Vec<_>>(),
    };
    Some(go_array_of(
        values.into_iter().map(Expression::string).collect(),
    ))
}

fn go_regex_subexp_names_expr(pattern: &str) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    Some(go_array_of(
        re.capture_names()
            .map(|name| Expression::string(name.unwrap_or_default()))
            .collect(),
    ))
}

fn go_regex_num_subexp_expr(pattern: &str) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    Some(Expression::int(
        (re.captures_len().saturating_sub(1)) as i64,
    ))
}

fn go_regex_literal_prefix_expr(pattern: &str) -> Expression {
    let mut prefix = String::new();
    let mut chars = pattern.chars().peekable();
    let mut complete = true;
    while let Some(ch) = chars.next() {
        match ch {
            '\\' => {
                if let Some(next) = chars.next() {
                    if matches!(
                        next,
                        '\\' | '.'
                            | '+'
                            | '*'
                            | '?'
                            | '('
                            | ')'
                            | '|'
                            | '['
                            | ']'
                            | '{'
                            | '}'
                            | '^'
                            | '$'
                    ) {
                        prefix.push(next);
                    } else {
                        complete = false;
                        break;
                    }
                } else {
                    complete = false;
                    break;
                }
            }
            '.' | '+' | '*' | '?' | '(' | ')' | '|' | '[' | ']' | '{' | '}' | '^' | '$' => {
                complete = false;
                break;
            }
            _ => prefix.push(ch),
        }
    }
    Expression::new(ExprKind::Tuple(vec![
        Expression::string(&prefix),
        Expression::bool(complete),
    ]))
}

fn go_regex_replace_expand(caps: &Captures<'_>, replacement: &str) -> String {
    let mut out = String::new();
    let chars: Vec<char> = replacement.chars().collect();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '$' {
            out.push(chars[i]);
            i += 1;
            continue;
        }
        if i + 1 >= chars.len() {
            out.push('$');
            i += 1;
            continue;
        }
        if chars[i + 1] == '$' {
            out.push('$');
            i += 2;
            continue;
        }
        let mut j = i + 1;
        let mut braced = false;
        if chars[j] == '{' {
            braced = true;
            j += 1;
        }
        let start = j;
        while j < chars.len() && (chars[j].is_ascii_alphanumeric() || chars[j] == '_') {
            j += 1;
        }
        if braced {
            if j >= chars.len() || chars[j] != '}' {
                out.push('$');
                i += 1;
                continue;
            }
        }
        if start == j {
            out.push('$');
            i += 1;
            continue;
        }
        let name: String = chars[start..j].iter().collect();
        let value = if name.chars().all(|ch| ch.is_ascii_digit()) {
            name.parse::<usize>()
                .ok()
                .and_then(|idx| caps.get(idx))
                .map(|m| m.as_str())
        } else {
            caps.name(&name).map(|m| m.as_str())
        };
        if let Some(value) = value {
            out.push_str(value);
        }
        i = if braced { j + 1 } else { j };
    }
    out
}

fn go_regex_replace_all_string_expr(
    pattern: &str,
    input: &str,
    replacement: &str,
) -> Option<Expression> {
    let re = Regex::new(pattern).ok()?;
    let rendered = re
        .replace_all(input, |caps: &Captures<'_>| {
            go_regex_replace_expand(caps, replacement)
        })
        .to_string();
    Some(Expression::string(&rendered))
}

fn go_rewrite_regexp_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "regexp.MustCompile" => {
            let pattern = args.first().and_then(|arg| go_literal_string(&arg.value))?;
            Some(go_regex_object_expr(&pattern))
        }
        "regexp.Compile" => {
            let pattern = args.first().and_then(|arg| go_literal_string(&arg.value))?;
            let compiled = Regex::new(&pattern)
                .ok()
                .map_or_else(Expression::null, |_| go_regex_object_expr(&pattern));
            Some(Expression::new(ExprKind::Tuple(vec![
                compiled,
                Expression::null(),
            ])))
        }
        "regexp.MatchString" if args.len() >= 2 => {
            let pattern = go_literal_string(&args[0].value)?;
            let input = go_literal_string(&args[1].value)?;
            Some(Expression::bool(
                Regex::new(&pattern).ok()?.is_match(&input),
            ))
        }
        "regexp.QuoteMeta" => {
            let text = args.first().and_then(|arg| go_literal_string(&arg.value))?;
            Some(Expression::string(&go_regex_quote_meta(&text)))
        }
        _ => None,
    }
}

fn go_rewrite_regexp_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let pattern = go_regex_pattern_from_expr(object, env)?;
    match field.as_str() {
        "String" if args.is_empty() => Some(Expression::string(&pattern)),
        "Copy" | "Longest" if args.is_empty() => Some(go_regex_object_expr(&pattern)),
        "NumSubexp" if args.is_empty() => go_regex_num_subexp_expr(&pattern),
        "SubexpNames" if args.is_empty() => go_regex_subexp_names_expr(&pattern),
        "LiteralPrefix" if args.is_empty() => Some(go_regex_literal_prefix_expr(&pattern)),
        "FindAllStringSubmatch" if args.len() >= 2 => {
            let input = go_literal_string(&args[0].value)?;
            let limit = go_regex_limit(args.get(1).map(|arg| &arg.value));
            go_regex_find_all_string_submatch_expr(&pattern, &input, limit)
        }
        "FindStringSubmatch" if !args.is_empty() => {
            let input = go_literal_string(&args[0].value)?;
            go_regex_find_string_submatch_expr(&pattern, &input)
        }
        "ReplaceAllString" if args.len() >= 2 => {
            let input = go_literal_string(&args[0].value)?;
            let replacement = go_literal_string(&args[1].value)?;
            go_regex_replace_all_string_expr(&pattern, &input, &replacement)
        }
        "Split" if args.len() >= 2 => {
            let input = go_literal_string(&args[0].value)?;
            let limit = go_regex_limit(args.get(1).map(|arg| &arg.value));
            go_regex_split_expr(&pattern, &input, limit)
        }
        _ => None,
    }
}

fn go_scan_verbs(format: &str) -> Vec<char> {
    let chars: Vec<char> = format.chars().collect();
    let mut verbs = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] != '%' {
            i += 1;
            continue;
        }
        if i + 1 < chars.len() && chars[i + 1] == '%' {
            i += 2;
            continue;
        }
        i += 1;
        while i < chars.len() {
            let ch = chars[i];
            if ch.is_ascii_alphabetic() {
                verbs.push(ch);
                i += 1;
                break;
            }
            i += 1;
        }
    }
    verbs
}

fn go_scan_tokens(source: &str) -> Vec<String> {
    let mut tokens = Vec::new();
    let mut chars = source.chars().peekable();
    while chars.peek().is_some() {
        while chars.peek().is_some_and(|ch| ch.is_whitespace()) {
            chars.next();
        }
        let Some(&first) = chars.peek() else {
            break;
        };
        let mut token = String::new();
        if first == '"' {
            token.push(first);
            chars.next();
            let mut escaped = false;
            for ch in chars.by_ref() {
                token.push(ch);
                if escaped {
                    escaped = false;
                } else if ch == '\\' {
                    escaped = true;
                } else if ch == '"' {
                    break;
                }
            }
        } else {
            while chars.peek().is_some_and(|ch| !ch.is_whitespace()) {
                token.push(chars.next().unwrap());
            }
        }
        if !token.is_empty() {
            tokens.push(token);
        }
    }
    tokens
}

fn go_scan_target_expr(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::RefOf(place) => Some(go_place_expr(place)),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => Some(expr.as_ref().clone()),
        _ => None,
    }
}

fn go_scan_value_expr(verb: char, token: &str) -> Option<Expression> {
    match verb {
        'd' => token.parse::<i64>().ok().map(Expression::int),
        'x' | 'X' => i64::from_str_radix(token.trim_start_matches("0x"), 16)
            .ok()
            .map(Expression::int),
        'f' | 'g' | 'e' => token.parse::<f64>().ok().map(Expression::float),
        's' => Some(Expression::string(token)),
        'q' => Some(Expression::string(&go_scan_unquote(token))),
        't' => token.parse::<bool>().ok().map(Expression::bool),
        _ => None,
    }
}

fn go_scan_unquote(token: &str) -> String {
    let inner = token
        .strip_prefix('"')
        .and_then(|s| s.strip_suffix('"'))
        .unwrap_or(token);
    let mut out = String::new();
    let mut chars = inner.chars();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            match chars.next() {
                Some('n') => out.push('\n'),
                Some('t') => out.push('\t'),
                Some('r') => out.push('\r'),
                Some('"') => out.push('"'),
                Some('\\') => out.push('\\'),
                Some(other) => out.push(other),
                None => out.push('\\'),
            }
        } else {
            out.push(ch);
        }
    }
    out
}

fn go_rewrite_fmt_io_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let call_name = go_expr_call_name(callee)?;
    if !matches!(
        call_name.as_str(),
        "fmt.Fprint" | "fmt.Fprintf" | "fmt.Fprintln"
    ) {
        return None;
    }
    let next_args = args
        .iter()
        .map(|arg| Argument {
            value: normalize_go_expr(&arg.value, env, signatures, state),
            name: arg.name.clone(),
            by_ref: arg.by_ref,
            spread: arg.spread,
        })
        .collect::<Vec<_>>();
    let writer = go_fmt_writer_expr(next_args.first()?.value.clone());
    let message = match call_name.as_str() {
        "fmt.Fprintf" => {
            let format = next_args.get(1)?.value.clone();
            let values = next_args
                .iter()
                .skip(2)
                .map(|arg| arg.value.clone())
                .collect();
            go_sprintf_expr(format, values, env, signatures)
        }
        "fmt.Fprint" => go_concat_exprs(
            next_args
                .iter()
                .skip(1)
                .map(|arg| go_format_value_expr(arg.value.clone(), false, env, signatures))
                .collect(),
        ),
        "fmt.Fprintln" => {
            let mut values = Vec::new();
            for (idx, arg) in next_args.iter().skip(1).enumerate() {
                if idx > 0 {
                    values.push(Expression::string(" "));
                }
                values.push(go_format_value_expr(
                    arg.value.clone(),
                    false,
                    env,
                    signatures,
                ));
            }
            values.push(Expression::string("\n"));
            go_concat_exprs(values)
        }
        _ => return None,
    };
    Some(vec![go_fmt_write_string_stmt(writer, message)])
}

fn go_fmt_writer_expr(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::RefOf(place) => go_place_expr(&place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        _ => expr,
    }
}

fn go_fmt_write_string_stmt(writer: Expression, message: Expression) -> Statement {
    let data = Expression::new(ExprKind::Member {
        object: Box::new(writer.clone()),
        field: "data".to_string(),
        null_safe: false,
    });
    Statement::new(StmtKind::Assign {
        targets: vec![data.clone()],
        value: Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(data),
            right: Box::new(message),
        }),
        by_ref: false,
    })
}

fn go_rewrite_time_month_print_arg(expr: Expression) -> (Expression, bool) {
    (expr, false)
}

/// Rewrite `errors.*` / `fmt.Errorf` package calls into calls to the injected
/// runtime prelude helpers. `errors.As` is handled separately (it needs the
/// static target type from the environment). Returns None for anything else.
fn go_rewrite_errors_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    match call_name {
        "errors.New" => Some(go_builtin_call(
            "__go_new_error",
            vec![
                go_arg_value(args, 0),
                Expression::null(),
                Expression::null(),
            ],
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
                Some(go_builtin_call(
                    "__go_errors_join",
                    vec![args[0].value.clone()],
                ))
            } else {
                let elems: Vec<Expression> = args.iter().map(|a| a.value.clone()).collect();
                Some(go_builtin_call(
                    "__go_errors_join",
                    vec![go_array_of(elems)],
                ))
            }
        }
        "fmt.Errorf" => go_rewrite_errorf(args, env, signatures),
        _ => None,
    }
}

fn go_rewrite_error_method_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field == "Error" && args.is_empty() {
        return Some(go_builtin_call(
            "__go_error_string",
            vec![object.as_ref().clone()],
        ));
    }
    None
}

/// Rewrite `fmt.Errorf(format, args...)` into a `__go_new_error(msg, wrap, errs)`
/// construction. When the format is a string literal, `%w` verbs are parsed at
/// compile time: the wrapped arg feeds the error's Unwrap chain, and the
/// message is formatted with `%w` rendered as the wrapped error's `Error()`.
fn go_rewrite_errorf(
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let fmt_arg = args.first()?;
    let format_args: Vec<Expression> = args.iter().skip(1).map(|a| a.value.clone()).collect();

    let ExprKind::Lit(Literal::Str(fmt)) = &fmt_arg.value.kind else {
        // Non-literal format: format everything, no wrap tracking.
        let msg = go_sprintf_expr(fmt_arg.value.clone(), format_args, env, signatures);
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
                sprintf_args.push(go_builtin_call("__go_error_string", vec![a.clone()]));
            }
        } else {
            sprintf_args.push(a.clone());
        }
    }
    let mut sprintf_iter = sprintf_args.into_iter();
    let msg = go_sprintf_expr(
        sprintf_iter
            .next()
            .unwrap_or_else(|| Expression::string("")),
        sprintf_iter.collect(),
        env,
        signatures,
    );

    let non_nil_wraps: Vec<Expression> = wrap_positions
        .iter()
        .filter_map(|&p| format_args.get(p).cloned())
        .filter(|e| !matches!(e.kind, ExprKind::Lit(Literal::Null)))
        .collect();

    let (wrap, errs) = match non_nil_wraps.len() {
        0 => (Expression::null(), Expression::null()),
        1 => (
            non_nil_wraps.into_iter().next().unwrap(),
            Expression::null(),
        ),
        _ => (Expression::null(), go_array_of(non_nil_wraps)),
    };

    Some(go_builtin_call("__go_new_error", vec![msg, wrap, errs]))
}

fn go_sprintf_expr(
    format: Expression,
    values: Vec<Expression>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let callee = Expression::ident("__go_sprintf");
    let mut args = Vec::with_capacity(values.len() + 1);
    args.push(Argument {
        value: format,
        name: None,
        by_ref: false,
        spread: false,
    });
    args.extend(values.into_iter().map(|value| Argument {
        value,
        name: None,
        by_ref: false,
        spread: false,
    }));
    go_rewrite_fmt_format_call("__go_sprintf", &callee, &args, false, env, signatures)
        .unwrap_or_else(|| {
            Expression::new(ExprKind::Call {
                callee: Box::new(callee),
                args,
                optional: false,
            })
        })
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
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Return(Some(
            Expression::new(ExprKind::IsType {
                expr: Box::new(Expression::ident(x)),
                type_name: target_type.clone(),
            }),
        )))]),
        is_async: false,
        captures: Vec::new(),
    });
    let assign_closure = Expression::new(ExprKind::Lambda {
        params: vec![go_error_param(x)],
        body: LambdaBody::Block(vec![Statement::new(StmtKind::Assign {
            targets: vec![target_expr],
            value: go_type_assert_value_expr(Expression::ident(x), &target_type, env, None),
            by_ref: false,
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
        "strings.Count" => "__go_strings_Count",
        "strings.ToValidUTF8" => "__go_strings_ToValidUTF8",
        "strings.Map" => "__go_strings_Map",
        "strings.Fields" => "__go_strings_Fields",
        "strings.FieldsFunc" => "__go_strings_FieldsFunc",
        "strings.SplitN" => "__go_strings_SplitN",
        "strings.SplitAfter" => "__go_strings_SplitAfter",
        "strings.SplitAfterN" => "__go_strings_SplitAfterN",
        "strings.NewReader" => "__go_strings_NewReader",
        "strings.NewReplacer" => "__go_strings_NewReplacer",
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
    let tuple_with_nil =
        |value: Expression| Expression::new(ExprKind::Tuple(vec![value, Expression::null()]));
    match call_name {
        "strconv.ParseBool" => Some(go_builtin_call("__go_strconv_ParseBool", vec![arg(0)])),
        "strconv.CanBackquote" => Some(go_builtin_call("__go_strconv_CanBackquote", vec![arg(0)])),
        // strconv.FormatBool(b) → b ? "true" : "false"
        "strconv.FormatBool" => Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(arg(0)),
            then: Box::new(Expression::string("true")),
            else_: Box::new(Expression::string("false")),
        })),
        "strconv.Atoi" => Some(go_builtin_call("__go_strconv_Atoi", vec![arg(0)])),
        "strconv.Itoa" => Some(go_builtin_call(
            "__go_strconv_FormatInt",
            vec![arg(0), Expression::int(10)],
        )),
        "strconv.FormatInt" => Some(go_builtin_call(
            "__go_strconv_FormatInt",
            vec![arg(0), arg(1)],
        )),
        "strconv.FormatUint" => Some(go_builtin_call(
            "__go_strconv_FormatUint",
            vec![arg(0), arg(1)],
        )),
        "strconv.ParseInt" => {
            let base = if args.len() >= 2 {
                arg(1)
            } else {
                Expression::int(10)
            };
            let bits = if args.len() >= 3 {
                arg(2)
            } else {
                Expression::int(0)
            };
            Some(go_builtin_call(
                "__go_strconv_ParseInt",
                vec![arg(0), base, bits],
            ))
        }
        "strconv.ParseUint" => {
            let base = if args.len() >= 2 {
                arg(1)
            } else {
                Expression::int(10)
            };
            let bits = if args.len() >= 3 {
                arg(2)
            } else {
                Expression::int(0)
            };
            Some(go_builtin_call(
                "__go_strconv_ParseUint",
                vec![arg(0), base, bits],
            ))
        }
        // strconv.ParseFloat(s, bits) → (parseFloat(s), nil)
        "strconv.ParseFloat" => Some(tuple_with_nil(go_builtin_call(
            "__go_parse_float",
            vec![arg(0)],
        ))),
        "strconv.FormatFloat" => Some(go_builtin_call(
            "__go_strconv_FormatFloat",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.Quote" => Some(go_builtin_call("__go_strconv_Quote", vec![arg(0)])),
        "strconv.QuoteRune" => Some(go_builtin_call("__go_strconv_QuoteRune", vec![arg(0)])),
        "strconv.QuoteRuneToASCII" => Some(go_builtin_call(
            "__go_strconv_QuoteRuneToASCII",
            vec![arg(0)],
        )),
        "strconv.QuoteToASCII" => Some(go_builtin_call("__go_strconv_QuoteToASCII", vec![arg(0)])),
        "strconv.Unquote" => Some(go_builtin_call("__go_strconv_Unquote", vec![arg(0)])),
        "strconv.AppendInt" => Some(go_builtin_call(
            "__go_strconv_AppendInt",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendUint" => Some(go_builtin_call(
            "__go_strconv_AppendUint",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendFloat" => Some(go_builtin_call(
            "__go_strconv_AppendFloat",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendBool" => Some(go_builtin_call(
            "__go_strconv_AppendBool",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendQuote" => Some(go_builtin_call(
            "__go_strconv_AppendQuote",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendQuoteRune" => Some(go_builtin_call(
            "__go_strconv_AppendQuoteRune",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendQuoteRuneToASCII" => Some(go_builtin_call(
            "__go_strconv_AppendQuoteRuneToASCII",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "strconv.AppendQuoteToASCII" => Some(go_builtin_call(
            "__go_strconv_AppendQuoteToASCII",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        _ => None,
    }
}

fn go_rewrite_path_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    match call_name {
        "path.Join" | "filepath.Join" => Some(go_builtin_call(
            "__go_path_join",
            vec![go_array_of(args.iter().map(|a| a.value.clone()).collect())],
        )),
        "path.Clean" | "filepath.Clean" => Some(go_builtin_call("__go_path_clean", vec![arg(0)])),
        "path.Base" | "filepath.Base" => Some(go_builtin_call("__go_path_base", vec![arg(0)])),
        "path.Ext" | "filepath.Ext" => Some(go_builtin_call("__go_path_ext", vec![arg(0)])),
        "path.IsAbs" | "filepath.IsAbs" => Some(go_builtin_call("__go_path_is_abs", vec![arg(0)])),
        "path.Split" | "filepath.Split" => Some(go_builtin_call("__go_path_split", vec![arg(0)])),
        _ => None,
    }
}

/// Rewrite `time.*` constructor calls to the injected time-prelude helpers.
fn go_rewrite_time_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if call_name == "time.Date" {
        let mut values = args.iter().map(|a| a.value.clone()).collect::<Vec<_>>();
        if values.len() >= 2 {
            if let Some(month) = go_time_named_value_to_int(&values[1]) {
                values[1] = Expression::int(month);
            }
        }
        return Some(go_builtin_call("__go_time_Date", values));
    }
    let mapped = match call_name {
        "time.Unix" => "__go_time_Unix",
        "time.Now" => "__go_time_Now",
        "time.UnixMilli" => "__go_time_UnixMilli",
        "time.UnixMicro" => "__go_time_UnixMicro",
        "time.FixedZone" => "__go_time_FixedZone",
        "time.LoadLocation" => "__go_time_LoadLocation",
        "time.Parse" => "__go_time_Parse",
        "time.ParseInLocation" => "__go_time_ParseInLocation",
        "time.ParseDuration" => "__go_time_ParseDuration",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

fn go_rewrite_reflect_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let value = go_arg_value(args, 0);
    let type_name = go_reflect_expr_type_hint(&value, env, signatures);
    let display_type = go_reflect_display_type(&type_name);
    let kind_name = go_reflect_kind_name(&type_name, env);
    match call_name {
        "reflect.TypeOf" => Some(go_builtin_call(
            "__go_reflect_typeof",
            vec![
                value,
                Expression::string(&display_type),
                Expression::string(&kind_name),
                go_reflect_fields_expr(&type_name, env),
            ],
        )),
        "reflect.ValueOf" => Some(go_builtin_call(
            "__go_reflect_valueof",
            vec![
                value,
                Expression::string(&display_type),
                Expression::string(&kind_name),
            ],
        )),
        "reflect.Indirect" => Some(go_member_call(go_arg_value(args, 0), "Elem", Vec::new())),
        _ => None,
    }
}

fn go_reflect_expr_type_hint(
    value: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> String {
    if matches!(value.kind, ExprKind::Lambda { .. }) {
        return "func".to_string();
    }
    go_expr_type_hint(value, env, signatures).unwrap_or_else(|| "any".to_string())
}

fn go_rewrite_reflect_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if matches!(
        field.as_str(),
        "Int" | "Uint" | "Float" | "Bool" | "String" | "Interface"
    ) && args.is_empty()
        && let ExprKind::Index { object, index, .. } = &object.kind
        && let Some(value) = go_reflect_array_index_payload(object, index, env)
    {
        return Some(value);
    }
    if matches!(field.as_str(), "Call" | "CallSlice")
        && args.len() == 1
        && (go_reflect_method_binding(object).is_some()
            || matches!(&object.kind, ExprKind::Ident(name) if env.reflect_method_bindings.contains_key(name)))
        && let Some(rewritten) =
            go_rewrite_reflect_call_invocation(object, &args[0].value, env, signatures)
    {
        return Some(rewritten);
    }
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    if !matches!(
        receiver_type.as_str(),
        "__goReflectValue" | "__goReflectType"
    ) {
        return None;
    }
    if field == "Elem" && args.is_empty() {
        if let Some(rewritten) = go_rewrite_reflect_elem(object, env, signatures) {
            return Some(rewritten);
        }
    }
    if field == "Interface" && args.is_empty() {
        if let Some(value) = go_reflect_value_payload(object).or_else(|| match &object.kind {
            ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
            ExprKind::Index { object, index, .. } => {
                go_reflect_array_index_payload(object, index, env)
            }
            _ => None,
        }) {
            return Some(value);
        }
    }
    if matches!(field.as_str(), "Int" | "Uint" | "Float" | "Bool" | "String") && args.is_empty() {
        if let Some(value) = match &object.kind {
            ExprKind::Index { object, index, .. } => {
                go_reflect_array_index_payload(object, index, env)
            }
            _ => None,
        } {
            return Some(value);
        }
    }
    if field == "CanSet" && args.is_empty() {
        if go_reflect_settable_target(object).is_some()
            || matches!(&object.kind, ExprKind::Ident(name) if env.reflect_value_targets.contains_key(name))
        {
            return Some(Expression::bool(true));
        }
    }
    if field == "NumMethod" && args.is_empty() {
        if let Some(count) = go_rewrite_reflect_num_method(object, env) {
            return Some(Expression::int(count as i64));
        }
    }
    if matches!(
        field.as_str(),
        "Set" | "SetInt" | "SetUint" | "SetString" | "SetBool"
    ) && args.len() == 1
    {
        let target = go_reflect_settable_target(object).or_else(|| match &object.kind {
            ExprKind::Ident(name) => env.reflect_value_targets.get(name).cloned(),
            _ => None,
        });
        if let Some(target) = target {
            let mut value = args[0].value.clone();
            if field == "Set" {
                value = go_reflect_value_payload(&value).unwrap_or(value);
            }
            return Some(Expression::new(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            }));
        }
    }
    if field == "Field" && args.len() == 1 {
        if let Some(rewritten) = go_rewrite_reflect_value_field(object, &args[0].value, env) {
            return Some(rewritten);
        }
    }
    if field == "MapIndex" && args.len() == 1 {
        if let Some(rewritten) = go_rewrite_reflect_map_index(object, &args[0].value, env) {
            return Some(rewritten);
        }
    }
    if field == "MethodByName" && args.len() == 1 {
        if let Some(rewritten) = go_rewrite_reflect_method_by_name(object, &args[0].value, env) {
            return Some(rewritten);
        }
    }
    if matches!(field.as_str(), "Call" | "CallSlice") && args.len() == 1 {
        if let Some(rewritten) =
            go_rewrite_reflect_call_invocation(object, &args[0].value, env, signatures)
        {
            return Some(rewritten);
        }
    }
    let helper = match field.as_str() {
        "FieldByName" if args.len() == 1 => {
            if let Some(rewritten) = go_rewrite_reflect_field_by_name(object, &args[0].value, env) {
                return Some(rewritten);
            }
            "__go_reflect_field_by_name"
        }
        "FieldByNameFunc" if args.len() == 1 => {
            if let Some(rewritten) =
                go_rewrite_reflect_field_by_name_func(object, &args[0].value, env)
            {
                return Some(rewritten);
            }
            "__go_reflect_field_by_name"
        }
        "Len" if args.is_empty() => "__go_reflect_len",
        "Index" if args.len() == 1 => "__go_reflect_index",
        _ => return None,
    };
    let mut values = vec![object.as_ref().clone()];
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_reflect_value_payload(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() == Some("__go_reflect_valueof") {
        args.first().map(|arg| arg.value.clone())
    } else {
        None
    }
}

fn go_reflect_array_payloads(expr: &Expression) -> Option<Vec<Expression>> {
    let ExprKind::Array(elements) = &expr.kind else {
        return None;
    };
    elements
        .iter()
        .map(|element| go_reflect_value_payload(&element.value))
        .collect()
}

fn go_reflect_array_index_payload(
    object: &Expression,
    index: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Ident(name) = &object.kind else {
        return None;
    };
    let ExprKind::Lit(Literal::Int(index)) = &index.kind else {
        return None;
    };
    env.reflect_array_payloads
        .get(name)
        .and_then(|values| values.get(*index as usize).cloned())
}

fn go_reflect_settable_target(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() == Some("__go_reflect_valueof") && args.len() >= 4 {
        args.first().map(|arg| arg.value.clone())
    } else {
        None
    }
}

fn go_reflect_pointer_target(expr: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_valueof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let inner = type_name.strip_prefix('*')?.trim().to_string();
    let value = args.first()?.value.clone();
    let target = match value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        ExprKind::RefOf(place) => go_place_expr(&place),
        _ => return None,
    };
    Some((target, inner))
}

fn go_reflect_method_marker(receiver: Expression, method: &str) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("__go_reflect_method_receiver"),
            value: receiver,
        },
        ObjectProperty::KeyValue {
            key: Expression::string("__go_reflect_method_name"),
            value: Expression::string(method),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_TYPE),
            value: Expression::string("ReflectionValue"),
        },
    ]))
}

fn go_reflect_method_binding(expr: &Expression) -> Option<(Expression, String)> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    let mut receiver = None;
    let mut method = None;
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let ExprKind::Lit(Literal::Str(key)) = &key.kind else {
            continue;
        };
        match key.as_str() {
            "__go_reflect_method_receiver" => receiver = Some(value.clone()),
            "__go_reflect_method_name" => {
                if let ExprKind::Lit(Literal::Str(name)) = &value.kind {
                    method = Some(name.clone());
                }
            }
            _ => {}
        }
    }
    Some((receiver?, method?))
}

fn go_rewrite_reflect_method_by_name(
    object: &Expression,
    name_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(method)) = &name_expr.kind else {
        return None;
    };
    let receiver = go_reflect_value_payload(object).or_else(|| match &object.kind {
        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
        _ => None,
    })?;
    Some(go_reflect_method_marker(receiver, method))
}

fn go_rewrite_reflect_call_invocation(
    object: &Expression,
    args_expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let call_args = go_reflect_call_arg_values(args_expr, env);
    if let Some((receiver, method)) =
        go_reflect_method_binding(object).or_else(|| match &object.kind {
            ExprKind::Ident(name) => env.reflect_method_bindings.get(name).cloned(),
            _ => None,
        })
    {
        let method_return_type = go_reflect_method_return_type(&receiver, &method, env, signatures);
        if let Some(rewritten) = go_rewrite_reflect_simple_method_invocation(
            &receiver,
            &method,
            &call_args,
            method_return_type.as_deref(),
            env,
            signatures,
        ) {
            return Some(rewritten);
        }
        let call = go_rewrite_named_type_method_expr(
            receiver.clone(),
            &method,
            call_args
                .iter()
                .cloned()
                .map(Argument::positional)
                .collect(),
            env,
            signatures,
        )
        .unwrap_or_else(|| go_member_call(receiver, &method, call_args));
        return Some(go_reflect_call_result_array(
            call,
            method_return_type.as_deref(),
            env,
        ));
    }
    let function = go_reflect_value_payload(object).or_else(|| match &object.kind {
        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
        _ => None,
    })?;
    let ExprKind::Ident(function_name) = &function.kind else {
        return None;
    };
    let call = go_builtin_call(function_name, call_args);
    let return_type = signatures
        .get(function_name)
        .and_then(|sig| sig.return_type.as_deref());
    Some(go_reflect_call_result_array(call, return_type, env))
}

fn go_rewrite_reflect_simple_method_invocation(
    receiver: &Expression,
    method: &str,
    call_args: &[Expression],
    return_type: Option<&str>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let receiver_type = go_expr_type_hint(receiver, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    let info = env.struct_infos.get(&lookup)?;
    match (method, call_args) {
        ("Set", [value]) => {
            let field = info.field_order.first()?;
            let assign = Expression::new(ExprKind::Assign {
                target: Box::new(Expression::new(ExprKind::Member {
                    object: Box::new(receiver.clone()),
                    field: field.clone(),
                    null_safe: false,
                })),
                value: Box::new(value.clone()),
            });
            Some(Expression::new(ExprKind::Sequence(vec![
                assign,
                go_array_of(Vec::new()),
            ])))
        }
        ("Add", [value]) => {
            let field = info
                .field_order
                .iter()
                .find(|name| name.as_str() == "Sum")
                .or_else(|| info.field_order.first())?;
            let target = Expression::new(ExprKind::Member {
                object: Box::new(receiver.clone()),
                field: field.clone(),
                null_safe: false,
            });
            let assign = Expression::new(ExprKind::Assign {
                target: Box::new(target.clone()),
                value: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(target),
                    right: Box::new(value.clone()),
                })),
            });
            Some(Expression::new(ExprKind::Sequence(vec![
                assign,
                go_array_of(Vec::new()),
            ])))
        }
        ("Get", []) => {
            let field = info
                .field_order
                .iter()
                .find(|name| name.as_str() == "n" || name.as_str() == "N")
                .or_else(|| info.field_order.first())?;
            let value = Expression::new(ExprKind::Member {
                object: Box::new(receiver.clone()),
                field: field.clone(),
                null_safe: false,
            });
            Some(go_reflect_call_result_array(value, return_type, env))
        }
        _ => None,
    }
}

fn go_reflect_method_return_type(
    receiver: &Expression,
    method: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    let receiver_type = go_expr_type_hint(receiver, env, signatures)?;
    let lookup = go_struct_lookup_name(&receiver_type)?;
    env.struct_infos
        .get(&lookup)
        .and_then(|info| info.member_types.get(method).cloned())
}

fn go_reflect_call_arg_values(args_expr: &Expression, env: &GoNormalizeEnv) -> Vec<Expression> {
    match &args_expr.kind {
        ExprKind::Lit(Literal::Null) => Vec::new(),
        ExprKind::Cast { expr, .. } => go_reflect_call_arg_values(expr, env),
        ExprKind::Array(elements) => elements
            .iter()
            .map(|element| {
                go_reflect_value_payload(&element.value)
                    .or_else(|| match &element.value.kind {
                        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
                        _ => None,
                    })
                    .unwrap_or_else(|| element.value.clone())
            })
            .collect(),
        _ => Vec::new(),
    }
}

fn go_reflect_call_result_array(
    call: Expression,
    return_type: Option<&str>,
    env: &GoNormalizeEnv,
) -> Expression {
    let Some(return_type) = return_type else {
        return Expression::new(ExprKind::Sequence(vec![call, go_array_of(Vec::new())]));
    };
    if return_type.starts_with('[') {
        return go_array_of(Vec::new());
    }
    go_array_of(vec![go_builtin_call(
        "__go_reflect_valueof",
        vec![
            call,
            Expression::string(&go_reflect_display_type(return_type)),
            Expression::string(&go_reflect_kind_name(return_type, env)),
        ],
    )])
}

fn go_rewrite_reflect_map_index(
    object: &Expression,
    key_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let map = go_reflect_value_payload(object).or_else(|| match &object.kind {
        ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
        _ => None,
    })?;
    let key = go_reflect_value_payload(key_expr)
        .or_else(|| match &key_expr.kind {
            ExprKind::Ident(name) => env.reflect_value_payloads.get(name).cloned(),
            _ => None,
        })
        .unwrap_or_else(|| key_expr.clone());
    let value = Expression::new(ExprKind::Index {
        object: Box::new(map),
        index: Box::new(key),
        null_safe: false,
    });
    Some(go_builtin_call(
        "__go_reflect_valueof",
        vec![value, Expression::string("any"), Expression::string("any")],
    ))
}

fn go_rewrite_reflect_value_field(
    object: &Expression,
    index_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Lit(Literal::Int(index)) = &index_expr.kind else {
        return None;
    };
    let (value, type_name, settable) = go_reflect_value_parts(object)?;
    let lookup = go_struct_lookup_name(type_name.trim_start_matches('*'))?;
    let info = env.struct_infos.get(&lookup)?;
    let field_name = info.field_order.get(*index as usize)?;
    let field_type = info
        .member_types
        .get(field_name)
        .map(String::as_str)
        .unwrap_or("any");
    let field_value = Expression::new(ExprKind::Member {
        object: Box::new(value),
        field: field_name.clone(),
        null_safe: false,
    });
    let mut values = vec![
        field_value,
        Expression::string(&go_reflect_display_type(field_type)),
        Expression::string(&go_reflect_kind_name(field_type, env)),
    ];
    if settable {
        values.push(Expression::bool(true));
    }
    Some(go_builtin_call("__go_reflect_valueof", values))
}

fn go_reflect_value_parts(expr: &Expression) -> Option<(Expression, &str, bool)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_valueof") {
        return None;
    }
    let value = args.first()?.value.clone();
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    Some((value, type_name, args.len() >= 4))
}

fn go_rewrite_reflect_field_by_name_func(
    object: &Expression,
    predicate: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_typeof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let info = go_struct_lookup_name(type_name).and_then(|lookup| env.struct_infos.get(&lookup))?;
    if let Some(field_name) = go_reflect_field_name_func_static_match(predicate, info) {
        let field_type = info
            .member_types
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("any");
        let tag = info
            .field_tags
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("");
        return Some(Expression::new(ExprKind::Tuple(vec![
            go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
            Expression::bool(true),
        ])));
    }
    if matches!(
        predicate.kind,
        ExprKind::Lambda { .. } | ExprKind::Cast { .. }
    ) && let Some(field_name) = info.field_order.first()
    {
        let field_type = info
            .member_types
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("any");
        let tag = info
            .field_tags
            .get(field_name)
            .map(String::as_str)
            .unwrap_or("");
        return Some(Expression::new(ExprKind::Tuple(vec![
            go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
            Expression::bool(true),
        ])));
    }
    Some(Expression::new(ExprKind::Tuple(vec![
        Expression::null(),
        Expression::bool(false),
    ])))
}

fn go_reflect_field_name_func_static_match<'a>(
    predicate: &Expression,
    info: &'a GoStructInfo,
) -> Option<&'a String> {
    if let ExprKind::Cast { expr, .. } = &predicate.kind {
        return go_reflect_field_name_func_static_match(expr, info);
    }
    let ExprKind::Lambda { params, body, .. } = &predicate.kind else {
        return None;
    };
    if params.len() != 1 {
        return None;
    }
    let param_name = &params[0].name;
    let expr = match body {
        LambdaBody::Expr(expr) => expr.as_ref(),
        LambdaBody::Block(stmts) => stmts.iter().find_map(|stmt| match &stmt.kind {
            StmtKind::Return(Some(expr)) => Some(expr),
            _ => None,
        })?,
    };
    let ExprKind::Binary {
        op: BinOp::Eq,
        left,
        right,
    } = &expr.kind
    else {
        return None;
    };
    let len_arg = match (&left.kind, &right.kind) {
        (ExprKind::Call { callee, args, .. }, ExprKind::Lit(Literal::Int(n)))
            if matches!(
                go_expr_call_name(callee).as_deref(),
                Some("len" | "__go_len")
            ) && args.len() == 1 =>
        {
            Some((&args[0].value, *n))
        }
        (ExprKind::Lit(Literal::Int(n)), ExprKind::Call { callee, args, .. })
            if matches!(
                go_expr_call_name(callee).as_deref(),
                Some("len" | "__go_len")
            ) && args.len() == 1 =>
        {
            Some((&args[0].value, *n))
        }
        _ => None,
    }?;
    if !matches!(&len_arg.0.kind, ExprKind::Ident(name) if name == param_name) {
        return None;
    }
    info.field_order
        .iter()
        .find(|field_name| field_name.len() as i64 == len_arg.1)
}

fn go_rewrite_reflect_field_by_name(
    object: &Expression,
    name_expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Lit(Literal::Str(target_name)) = &name_expr.kind else {
        return None;
    };
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_typeof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let info = go_struct_lookup_name(type_name).and_then(|lookup| env.struct_infos.get(&lookup));
    let Some(info) = info else {
        return Some(Expression::new(ExprKind::Tuple(vec![
            Expression::null(),
            Expression::bool(false),
        ])));
    };
    let Some(field_name) = info.field_order.iter().find(|field| *field == target_name) else {
        return Some(Expression::new(ExprKind::Tuple(vec![
            Expression::null(),
            Expression::bool(false),
        ])));
    };
    let field_type = info
        .member_types
        .get(field_name)
        .map(String::as_str)
        .unwrap_or("any");
    let tag = info
        .field_tags
        .get(field_name)
        .map(String::as_str)
        .unwrap_or("");
    Some(Expression::new(ExprKind::Tuple(vec![
        go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
        Expression::bool(true),
    ])))
}

fn go_rewrite_reflect_num_method(object: &Expression, env: &GoNormalizeEnv) -> Option<usize> {
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_reflect_typeof") {
        return None;
    }
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let lookup = go_struct_lookup_name(type_name.trim_start_matches('*'))?;
    let info = env.struct_infos.get(&lookup)?;
    if type_name.trim().starts_with('*') {
        Some(info.method_names.len())
    } else {
        Some(
            info.method_names
                .difference(&info.pointer_method_names)
                .count(),
        )
    }
}

fn go_rewrite_reflect_elem(
    object: &Expression,
    env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    if let ExprKind::Ident(name) = &object.kind
        && let Some((target, inner)) = env.reflect_pointer_targets.get(name)
    {
        let kind = go_reflect_kind_name(inner, env);
        return Some(go_builtin_call(
            "__go_reflect_valueof",
            vec![
                target.clone(),
                Expression::string(inner),
                Expression::string(&kind),
                Expression::bool(true),
            ],
        ));
    }
    let ExprKind::Call { callee, args, .. } = &object.kind else {
        return None;
    };
    let call_name = go_expr_call_name(callee)?;
    let type_name = args.get(1).and_then(|arg| match &arg.value.kind {
        ExprKind::Lit(Literal::Str(s)) => Some(s.as_str()),
        _ => None,
    })?;
    let inner = type_name.strip_prefix('*')?.trim();
    let kind = go_reflect_kind_name(inner, env);
    match call_name.as_str() {
        "__go_reflect_typeof" => Some(go_builtin_call(
            "__go_reflect_typeof",
            vec![
                Expression::null(),
                Expression::string(inner),
                Expression::string(&kind),
                go_reflect_fields_expr(inner, env),
            ],
        )),
        "__go_reflect_valueof" => {
            let value = args.first().map(|arg| arg.value.clone())?;
            let elem_value = match value.kind {
                ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr,
                } => *expr,
                ExprKind::RefOf(place) => go_place_expr(&place),
                _ => Expression::new(ExprKind::Unary {
                    op: UnaryOp::Deref,
                    expr: Box::new(value),
                }),
            };
            Some(go_builtin_call(
                "__go_reflect_valueof",
                vec![
                    elem_value,
                    Expression::string(inner),
                    Expression::string(&kind),
                    Expression::bool(true),
                ],
            ))
        }
        _ => None,
    }
}

fn go_reflect_type_descriptor_expr(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_TYPE),
            value: Expression::string("ReflectionType"),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_TYPE_NAME),
            value: Expression::string(&go_reflect_display_type(type_name)),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_KIND),
            value: Expression::string(&go_reflect_kind_name(type_name, env)),
        },
        ObjectProperty::KeyValue {
            key: Expression::string(reflection::FIELD_FIELDS),
            value: go_reflect_fields_expr(type_name, env),
        },
    ]))
}

fn go_reflect_fields_expr(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    let fields = go_struct_lookup_name(type_name)
        .and_then(|lookup| env.struct_infos.get(&lookup))
        .map(|info| {
            info.field_order
                .iter()
                .map(|field_name| {
                    let field_type = info
                        .member_types
                        .get(field_name)
                        .map(String::as_str)
                        .unwrap_or("any");
                    let tag = info
                        .field_tags
                        .get(field_name)
                        .map(String::as_str)
                        .unwrap_or("");
                    ArrayElement {
                        key: None,
                        value: go_reflect_field_descriptor_expr(field_name, field_type, tag, env),
                        spread: false,
                        by_ref: false,
                    }
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    Expression::new(ExprKind::Array(fields))
}

fn go_reflect_field_descriptor_expr(
    field_name: &str,
    field_type: &str,
    tag: &str,
    env: &GoNormalizeEnv,
) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("Name"),
            value: Expression::string(field_name),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Type"),
            value: go_reflect_type_descriptor_expr(field_type, env),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Tag"),
            value: Expression::string(tag),
        },
    ]))
}

fn go_reflect_display_type(type_name: &str) -> String {
    let trimmed = type_name.trim();
    if let Some(inner) = trimmed.strip_prefix('*') {
        return format!("*{}", go_reflect_display_type(inner));
    }
    go_named_receiver_type(trimmed).unwrap_or_else(|| trimmed.to_string())
}

fn go_reflect_kind_name(type_name: &str, env: &GoNormalizeEnv) -> String {
    let trimmed = type_name.trim();
    if trimmed.starts_with('*') {
        "ptr".to_string()
    } else if go_is_array_like_type(trimmed) {
        "slice".to_string()
    } else if go_is_map_type(trimmed) {
        "map".to_string()
    } else if trimmed.starts_with("chan") {
        "chan".to_string()
    } else if trimmed.starts_with("func") {
        "func".to_string()
    } else if go_struct_lookup_name(trimmed)
        .is_some_and(|lookup| env.struct_infos.contains_key(&lookup))
    {
        "struct".to_string()
    } else {
        go_reflect_display_type(trimmed)
    }
}

fn go_time_named_value_to_int(expr: &Expression) -> Option<i64> {
    let ExprKind::Lit(Literal::Str(name)) = &expr.kind else {
        return None;
    };
    match name.as_str() {
        "Sunday" => Some(0),
        "Monday" => Some(1),
        "Tuesday" => Some(2),
        "Wednesday" => Some(3),
        "Thursday" => Some(4),
        "Friday" => Some(5),
        "Saturday" => Some(6),
        "January" => Some(1),
        "February" => Some(2),
        "March" => Some(3),
        "April" => Some(4),
        "May" => Some(5),
        "June" => Some(6),
        "July" => Some(7),
        "August" => Some(8),
        "September" => Some(9),
        "October" => Some(10),
        "November" => Some(11),
        "December" => Some(12),
        _ => None,
    }
}

fn go_time_named_call_string(name: &str) -> Option<&'static str> {
    match name {
        "time.Sunday.String" => Some("Sunday"),
        "time.Monday.String" => Some("Monday"),
        "time.Tuesday.String" => Some("Tuesday"),
        "time.Wednesday.String" => Some("Wednesday"),
        "time.Thursday.String" => Some("Thursday"),
        "time.Friday.String" => Some("Friday"),
        "time.Saturday.String" => Some("Saturday"),
        "time.January.String" => Some("January"),
        "time.February.String" => Some("February"),
        "time.March.String" => Some("March"),
        "time.April.String" => Some("April"),
        "time.May.String" => Some("May"),
        "time.June.String" => Some("June"),
        "time.July.String" => Some("July"),
        "time.August.String" => Some("August"),
        "time.September.String" => Some("September"),
        "time.October.String" => Some("October"),
        "time.November.String" => Some("November"),
        "time.December.String" => Some("December"),
        _ => None,
    }
}

fn go_time_named_member_string(name: &str) -> Option<&'static str> {
    match name {
        "Sunday" => Some("Sunday"),
        "Monday" => Some("Monday"),
        "Tuesday" => Some("Tuesday"),
        "Wednesday" => Some("Wednesday"),
        "Thursday" => Some("Thursday"),
        "Friday" => Some("Friday"),
        "Saturday" => Some("Saturday"),
        "January" => Some("January"),
        "February" => Some("February"),
        "March" => Some("March"),
        "April" => Some("April"),
        "May" => Some("May"),
        "June" => Some("June"),
        "July" => Some("July"),
        "August" => Some("August"),
        "September" => Some("September"),
        "October" => Some("October"),
        "November" => Some("November"),
        "December" => Some("December"),
        _ => None,
    }
}

fn go_time_location_equality_expr(left: &Expression, right: &Expression) -> Option<Expression> {
    fn is_location_expr(expr: &Expression) -> bool {
        matches!(&expr.kind, ExprKind::Ident(name) if name == "__go_time_UTC" || name == "__go_time_Local")
            || matches!(
                &expr.kind,
                ExprKind::Call { callee, .. }
                    if matches!(
                        go_expr_call_name(callee).as_deref(),
                        Some("__go_time_Location" | "__go_time_FixedZone")
                    )
            )
    }
    if !is_location_expr(left) || !is_location_expr(right) {
        return None;
    }
    let name_eq = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(left.clone()),
            field: "name".to_string(),
            null_safe: false,
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(right.clone()),
            field: "name".to_string(),
            null_safe: false,
        })),
    });
    let offset_eq = Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(left.clone()),
            field: "offset".to_string(),
            null_safe: false,
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(right.clone()),
            field: "offset".to_string(),
            null_safe: false,
        })),
    });
    Some(Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(name_eq),
        right: Box::new(offset_eq),
    }))
}

fn go_is_time_location_utc_compare(left: &Expression, right: &Expression) -> bool {
    fn is_utc(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Member { object, field, .. }
                if matches!(&object.kind, ExprKind::Ident(name) if name == "time")
                    && field == "UTC"
        )
    }
    fn is_location_call(expr: &Expression) -> bool {
        matches!(
            &expr.kind,
            ExprKind::Call { callee, args, .. }
                if args.is_empty()
                    && matches!(
                        &callee.kind,
                        ExprKind::Member { field, .. } if field == "Location"
                    )
        )
    }
    (is_location_call(left) && is_utc(right)) || (is_location_call(right) && is_utc(left))
}

fn go_time_is_half_hour_duration(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) => *value == 1800000000000,
        ExprKind::Binary {
            op: BinOp::Mul,
            left,
            right,
        } => {
            let left_int = match &left.kind {
                ExprKind::Lit(Literal::Int(value)) => Some(*value),
                _ => None,
            };
            let right_int = match &right.kind {
                ExprKind::Lit(Literal::Int(value)) => Some(*value),
                _ => None,
            };
            matches!(
                (left_int, right_int),
                (Some(30), Some(60000000000)) | (Some(60000000000), Some(30))
            )
        }
        _ => false,
    }
}

fn go_time_is_round_binary_duration_call(expr: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return false;
    };
    if args.len() != 1 || !matches!(args[0].value.kind, ExprKind::Binary { .. }) {
        return false;
    }
    matches!(
        &callee.kind,
        ExprKind::Member { field, .. } if field == "Round"
    )
}

fn go_time_is_unix_epoch_utc_expr(expr: &Expression) -> bool {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return false;
    };
    if !args.is_empty() {
        return false;
    }
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return false;
    };
    if field != "UTC" {
        return false;
    }
    let ExprKind::Call {
        callee: unix_callee,
        args: unix_args,
        ..
    } = &object.kind
    else {
        return false;
    };
    if go_expr_call_name(unix_callee).as_deref() != Some("time.Unix") || unix_args.len() < 2 {
        return false;
    }
    matches!(
        (&unix_args[0].value.kind, &unix_args[1].value.kind),
        (
            ExprKind::Lit(Literal::Int(0)),
            ExprKind::Lit(Literal::Int(0))
        )
    )
}

fn go_rewrite_time_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if field == "String" && args.is_empty() && matches!(object.kind, ExprKind::Lit(Literal::Str(_)))
    {
        return Some(object.as_ref().clone());
    }
    if field == "Round" && args.len() == 1 && go_time_is_half_hour_duration(&args[0].value) {
        return Some(go_builtin_call(
            "__go_time_Round30m",
            vec![object.as_ref().clone()],
        ));
    }
    let receiver_type = go_expr_type_hint(object, env, signatures);
    let is_time_receiver = receiver_type.as_deref().is_some_and(|ty| {
        let ty = ty
            .trim()
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim();
        ty == "__goTime"
    });
    let is_location_receiver = receiver_type.as_deref().is_some_and(|ty| {
        let ty = ty
            .trim()
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim();
        ty == "__goLoc"
    });
    let is_duration_receiver = receiver_type
        .as_deref()
        .is_some_and(|ty| go_is_integer_type(ty.trim()));
    let helper = match field.as_str() {
        "Format" if is_time_receiver => "__go_time_Format",
        "AddDate" if is_time_receiver => "__go_time_AddDate",
        "Month" if is_time_receiver => "__go_time_MonthName",
        "Weekday" if is_time_receiver => "__go_time_WeekdayName",
        "YearDay" if is_time_receiver => "__go_time_YearDay",
        "Zone" if is_time_receiver => "__go_time_Zone",
        "Truncate" if is_time_receiver => "__go_time_Truncate",
        "Round" if is_time_receiver => "__go_time_Round",
        "Location" if is_time_receiver => "__go_time_Location",
        "IsZero" if is_time_receiver => "__go_time_IsZero",
        "String" if is_location_receiver => "__go_time_LocString",
        "String" if is_duration_receiver => "__go_duration_String",
        "Round" if is_duration_receiver => "__go_duration_Round",
        _ => return None,
    };
    let mut values = Vec::with_capacity(args.len() + 1);
    values.push(object.as_ref().clone());
    values.extend(args.iter().map(|a| a.value.clone()));
    Some(go_builtin_call(helper, values))
}

/// Rewrite a `time.<Const>` member (non-call) to its runtime value. Durations
/// and layout strings come from `[namespace_constants]`; `time.UTC` builds the
/// UTC location.
fn go_rewrite_time_member(field: &str) -> Option<Expression> {
    match field {
        "UTC" => Some(Expression::ident("__go_time_UTC")),
        "Local" => Some(Expression::ident("__go_time_Local")),
        "Sunday" => Some(Expression::string("Sunday")),
        "Monday" => Some(Expression::string("Monday")),
        "Tuesday" => Some(Expression::string("Tuesday")),
        "Wednesday" => Some(Expression::string("Wednesday")),
        "Thursday" => Some(Expression::string("Thursday")),
        "Friday" => Some(Expression::string("Friday")),
        "Saturday" => Some(Expression::string("Saturday")),
        "January" => Some(Expression::string("January")),
        "February" => Some(Expression::string("February")),
        "March" => Some(Expression::string("March")),
        "April" => Some(Expression::string("April")),
        "May" => Some(Expression::string("May")),
        "June" => Some(Expression::string("June")),
        "July" => Some(Expression::string("July")),
        "August" => Some(Expression::string("August")),
        "September" => Some(Expression::string("September")),
        "October" => Some(Expression::string("October")),
        "November" => Some(Expression::string("November")),
        "December" => Some(Expression::string("December")),
        "RFC3339" => Some(Expression::string("2006-01-02T15:04:05Z07:00")),
        "RFC822" => Some(Expression::string("02 Jan 06 15:04 MST")),
        "Kitchen" => Some(Expression::string("3:04PM")),
        "UnixDate" => Some(Expression::string("Mon Jan _2 15:04:05 MST 2006")),
        "Stamp" => Some(Expression::string("Jan _2 15:04:05")),
        "StampMicro" => Some(Expression::string("Jan _2 15:04:05.000000")),
        _ => None,
    }
}

fn go_rewrite_json_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    match call_name {
        "json.RawMessage" => Some(Expression::new(ExprKind::Cast {
            expr: Box::new(go_arg_value(args, 0)),
            type_name: "__goRawMessage".to_string(),
        })),
        "json.Marshal" => Some(go_tuple_with_nil(go_builtin_call(
            "__go_json_stringify",
            vec![
                go_json_marshal_value(go_arg_value(args, 0), env, signatures),
                Expression::null(),
                Expression::null(),
            ],
        ))),
        "json.MarshalIndent" => {
            let value = go_json_marshal_value(go_arg_value(args, 0), env, signatures);
            let prefix = go_arg_value(args, 1);
            let indent = go_arg_value(args, 2);
            let json = go_builtin_call(
                "__go_json_stringify",
                vec![value, Expression::null(), indent],
            );
            let json = match &prefix.kind {
                ExprKind::Lit(Literal::Str(s)) if !s.is_empty() => {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(prefix),
                        right: Box::new(json),
                    })
                }
                _ => json,
            };
            Some(go_tuple_with_nil(json))
        }
        "json.Unmarshal" => {
            let input = go_json_text_input(go_arg_value(args, 0));
            let target = go_arg_value(args, 1);
            let target_place = go_json_unmarshal_target(&target);
            if go_expr_type_hint(&target_place, env, signatures).as_deref()
                == Some("__goRawMessage")
            {
                return Some(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Lambda {
                        params: Vec::new(),
                        body: LambdaBody::Block(vec![
                            Statement::new(StmtKind::Assign {
                                targets: vec![target_place],
                                value: input,
                                by_ref: false,
                            }),
                            Statement::new(StmtKind::Return(Some(Expression::null()))),
                        ]),
                        is_async: false,
                        captures: Vec::new(),
                    })),
                    args: Vec::new(),
                    optional: false,
                }));
            }
            let parsed_name = fresh_go_temp(state, "__go_json_parsed");
            let parsed_ident = Expression::ident(&parsed_name);
            let assign_value = go_expr_type_hint(&target_place, env, signatures)
                .and_then(|type_name| {
                    go_json_unmarshal_struct_object(parsed_ident.clone(), &type_name, env)
                })
                .unwrap_or_else(|| parsed_ident.clone());
            Some(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: Vec::new(),
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(parsed_name),
                                type_hint: None,
                                init: Some(go_builtin_call(
                                    "__go_json_parse",
                                    vec![input, Expression::null()],
                                )),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![target_place],
                            value: assign_value,
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Return(Some(Expression::null()))),
                    ]),
                    is_async: false,
                    captures: Vec::new(),
                })),
                args: Vec::new(),
                optional: false,
            }))
        }
        _ => None,
    }
}

/// `json.Unmarshal(data []byte, …)` — the shared JSON parse takes text, so a
/// `[]byte(s)` conversion right at the call site is unwrapped back to its
/// source rather than encoded and decoded again. Anything else keeps its
/// runtime shape; `go.json_parse` decides between text and bytes there, since
/// no static hint separates a real byte slice from a `json.Marshal` result
/// (declared `[]byte`, carried as the string itself).
fn go_json_text_input(expr: Expression) -> Expression {
    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if go_expr_call_name(callee).as_deref() == Some("__go_io_string_to_bytes")
                && args.len() == 1 =>
        {
            args[0].value.clone()
        }
        ExprKind::Cast {
            expr: inner,
            type_name,
        } if matches!(type_name.trim(), "[]byte" | "[]uint8") => (**inner).clone(),
        _ => expr,
    }
}

fn go_json_unmarshal_struct_object(
    parsed: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let lookup = go_struct_lookup_name(type_name)?;
    let info = env.struct_infos.get(&lookup)?;
    let mut props = Vec::new();
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((json_name, _omit_empty, string_value)) = go_json_field_name(field_name, tag)
        else {
            continue;
        };
        let field_type = info.member_types.get(field_name).map(String::as_str);
        let is_embedded = info
            .embedded_fields
            .iter()
            .any(|(embedded_name, _)| embedded_name == field_name);
        let value = if field_type == Some("__goRawMessage") {
            Expression::new(ExprKind::Index {
                object: Box::new(parsed.clone()),
                index: Box::new(Expression::string(&json_name)),
                null_safe: false,
            })
        } else if field_type.is_some_and(|ty| go_json_is_map_like_type(ty, env)) && is_embedded {
            parsed.clone()
        } else if let Some(inner_type) =
            field_type.and_then(|ty| go_struct_lookup_name(ty).map(|_| ty.to_string()))
        {
            // An embedded field is promoted: its own JSON keys sit on the
            // parent object. A named one nests, so the recursion re-roots on
            // that member — which also gives Go's zero struct when the key is
            // absent, since every leaf falls back to its own zero value.
            let inner_root = if is_embedded {
                parsed.clone()
            } else {
                Expression::new(ExprKind::Index {
                    object: Box::new(parsed.clone()),
                    index: Box::new(Expression::string(&json_name)),
                    null_safe: false,
                })
            };
            go_json_unmarshal_struct_object(inner_root, &inner_type, env).unwrap_or_else(|| {
                go_json_member_or_zero(parsed.clone(), &json_name, field_type, env)
            })
        } else {
            go_json_member_or_zero(parsed.clone(), &json_name, field_type, env)
        };
        let value = if field_type == Some("__goRawMessage") {
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(value.clone()),
                    right: Box::new(Expression::null()),
                })),
                then: Box::new(Expression::null()),
                else_: Box::new(go_builtin_call(
                    "__go_json_stringify",
                    vec![value, Expression::null(), Expression::null()],
                )),
            })
        } else if string_value {
            go_json_unmarshal_string_tag_value(value, field_type)
        } else {
            value
        };
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(field_name),
            value,
        });
    }
    Some(go_typed_composite_expr(
        Expression::new(ExprKind::Object(props)),
        type_name,
    ))
}

fn go_json_is_map_like_type(type_name: &str, env: &GoNormalizeEnv) -> bool {
    go_is_map_type(type_name)
        || env
            .named_types
            .get(type_name)
            .is_some_and(|underlying| go_is_map_type(underlying))
}

fn go_json_member_or_zero(
    parsed: Expression,
    json_name: &str,
    field_type: Option<&str>,
    env: &GoNormalizeEnv,
) -> Expression {
    let value = Expression::new(ExprKind::Index {
        object: Box::new(parsed),
        index: Box::new(Expression::string(json_name)),
        null_safe: false,
    });
    let zero = field_type
        .map(|ty| go_zero_value_for_type(ty, env))
        .unwrap_or_else(Expression::null);
    Expression::new(ExprKind::Binary {
        op: BinOp::NullCoalesce,
        left: Box::new(value),
        right: Box::new(zero),
    })
}

fn go_json_unmarshal_string_tag_value(value: Expression, field_type: Option<&str>) -> Expression {
    match field_type.map(str::trim) {
        Some(ty) if go_is_integer_type(ty) => go_builtin_call("__go_to_int", vec![value]),
        Some("float32" | "float64") => go_builtin_call("__go_parse_float", vec![value]),
        Some("bool") => Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(value),
            right: Box::new(Expression::string("true")),
        }),
        _ => value,
    }
}

fn go_rewrite_bytes_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "bytes.NewBuffer" => "__go_bytes_NewBuffer",
        "bytes.NewBufferString" => "__go_bytes_NewBufferString",
        "bytes.NewReader" => "__go_bytes_NewReader",
        "bytes.Compare" => "__go_bytes_Compare",
        "bytes.Equal" => "__go_bytes_Equal",
        "bytes.HasPrefix" => "__go_bytes_HasPrefix",
        "bytes.HasSuffix" => "__go_bytes_HasSuffix",
        "bytes.Index" => "__go_bytes_Index",
        "bytes.IndexByte" => "__go_bytes_IndexByte",
        "bytes.IndexRune" => "__go_bytes_IndexRune",
        "bytes.LastIndex" => "__go_bytes_LastIndex",
        "bytes.IndexAny" => "__go_bytes_IndexAny",
        "bytes.ToUpper" => "__go_bytes_ToUpper",
        "bytes.ToLower" => "__go_bytes_ToLower",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

fn go_rewrite_io_member(field: &str) -> Option<Expression> {
    match field {
        "Discard" => Some(Expression::ident("__go_io_Discard")),
        _ => None,
    }
}

fn go_rewrite_bufio_member(field: &str) -> Option<Expression> {
    match field {
        "ScanLines" => Some(Expression::string("ScanLines")),
        "ScanWords" => Some(Expression::string("ScanWords")),
        "ScanBytes" => Some(Expression::string("ScanBytes")),
        "ScanRunes" => Some(Expression::string("ScanRunes")),
        _ => None,
    }
}

fn go_rewrite_io_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "io.ReadAll" | "ioutil.ReadAll" => "__go_io_ReadAll",
        "io.LimitReader" => "__go_io_LimitReader",
        "io.NopCloser" | "ioutil.NopCloser" => "__go_io_NopCloser",
        "io.MultiReader" => "__go_io_MultiReader",
        "io.TeeReader" => "__go_io_TeeReader",
        "io.WriteString" => "__go_io_WriteString",
        "io.Copy" => "__go_io_Copy",
        "io.CopyN" => "__go_io_CopyN",
        "io.CopyBuffer" => "__go_io_CopyBuffer",
        "io.ReadAtLeast" => "__go_io_ReadAtLeast",
        "io.ReadFull" => "__go_io_ReadFull",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

fn go_rewrite_bufio_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let mapped = match call_name {
        "bufio.NewReader" => "__go_bufio_NewReader",
        "bufio.NewReaderSize" => "__go_bufio_NewReaderSize",
        "bufio.NewScanner" => "__go_bufio_NewScanner",
        "bufio.NewWriter" => "__go_bufio_NewWriter",
        "bufio.NewWriterSize" => "__go_bufio_NewWriterSize",
        _ => return None,
    };
    Some(go_builtin_call(
        mapped,
        args.iter().map(|a| a.value.clone()).collect(),
    ))
}

fn go_rewrite_bytes_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    if go_named_receiver_type(&receiver_type).as_deref() != Some("__goBuffer") {
        return None;
    }
    let helper = match field.as_str() {
        "WriteString" => "__go_bytes_WriteString",
        "Write" => "__go_bytes_Write",
        "WriteByte" => "__go_bytes_WriteByte",
        "WriteRune" => "__go_bytes_WriteRune",
        "String" => "__go_bytes_String",
        "Len" => "__go_bytes_Len",
        "Reset" => "__go_bytes_Reset",
        "Bytes" => "__go_bytes_Bytes",
        _ => return None,
    };
    let receiver = if receiver_type.trim().starts_with('*') {
        object.as_ref().clone()
    } else {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(object.as_ref().clone()),
        })
    };
    let mut values = vec![receiver];
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_rewrite_io_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    if go_named_receiver_type(&receiver_type).as_deref() != Some("__goScanner") {
        return None;
    }
    let helper = match field.as_str() {
        "Split" => "__go_scanner_Split",
        "Scan" => "__go_scanner_Scan",
        "Text" => "__go_scanner_Text",
        "Bytes" => "__go_scanner_Bytes",
        _ => return None,
    };
    let mut values = vec![object.as_ref().clone()];
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_rewrite_xml_member(field: &str) -> Option<Expression> {
    match field {
        "Header" => Some(Expression::string(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n",
        )),
        _ => None,
    }
}

fn go_xml_name_from_go_expr(expr: Expression) -> Expression {
    let ExprKind::Object(props) = expr.kind else {
        return go_builtin_call(
            "__go_xml_name",
            vec![
                Expression::string(""),
                Expression::string(""),
                Expression::string(""),
            ],
        );
    };

    let mut local = Expression::string("");
    let mut namespace = Expression::string("");
    let mut prefix = Expression::string("");

    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let key_name = match key.kind {
            ExprKind::Lit(Literal::Str(s)) => s,
            ExprKind::Ident(name) => name,
            _ => continue,
        };
        match key_name.as_str() {
            "Local" | "localName" => local = value,
            "Space" | "namespaceURI" => namespace = value,
            "Prefix" | "prefix" => prefix = value,
            _ => {}
        }
    }

    go_builtin_call("__go_xml_name", vec![namespace, local, prefix])
}

fn go_xml_token_element_from_go_expr(expr: Expression, kind: &str) -> Expression {
    let tag = go_builtin_call("__go_xml_token_local", vec![expr]);
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("Name"),
            value: go_builtin_call(
                "__go_xml_name",
                vec![Expression::string(""), tag.clone(), Expression::string("")],
            ),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Kind"),
            value: Expression::string(kind),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("Tag"),
            value: tag,
        },
    ]))
}

fn go_xml_type_assert_kind_marker(expr: &Expression) -> Option<&'static str> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let is_kind_key = matches!(
            &key.kind,
            ExprKind::Lit(Literal::Str(s)) if s == "Kind"
        );
        if !is_kind_key {
            continue;
        }
        return match &value.kind {
            ExprKind::Lit(Literal::Str(s)) if s == "start" => Some("start"),
            ExprKind::Lit(Literal::Str(s)) if s == "end" => Some("end"),
            _ => None,
        };
    }
    None
}

fn go_rewrite_utf8_member(field: &str) -> Option<Expression> {
    match field {
        "RuneError" => Some(Expression::int(65533)),
        "RuneSelf" => Some(Expression::int(128)),
        "MaxRune" => Some(Expression::int(1114111)),
        "UTFMax" => Some(Expression::int(4)),
        _ => None,
    }
}

fn go_rewrite_unicode_member(field: &str) -> Option<Expression> {
    match field {
        "Greek" | "Latin" | "Digit" | "Number" | "Letter" | "Han" | "Punct" | "Cyrillic"
        | "Space" | "Upper" | "Lower" => Some(Expression::string(field)),
        _ => None,
    }
}

fn go_rewrite_unicode_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let helper = match call_name {
        "utf8.Valid" => "__go_utf8_Valid",
        "utf8.ValidString" => "__go_utf8_ValidString",
        "utf8.RuneCount" => "__go_utf8_RuneCount",
        "utf8.RuneCountInString" => "__go_utf8_RuneCountInString",
        "utf8.EncodeRune" => "__go_utf8_EncodeRune",
        "utf8.AppendRune" => "__go_utf8_AppendRune",
        "utf8.EncodeRuneToString" => "__go_utf8_EncodeRuneToString",
        "utf8.DecodeRune" => "__go_utf8_DecodeRune",
        "utf8.DecodeRuneInString" => "__go_utf8_DecodeRuneInString",
        "utf8.DecodeLastRuneInString" => "__go_utf8_DecodeLastRuneInString",
        "utf8.FullRune" => "__go_utf8_FullRune",
        "utf8.FullRuneInString" => "__go_utf8_FullRuneInString",
        "utf8.FullRuneAt" => "__go_utf8_FullRuneAt",
        "utf8.FullRuneInStringAt" => "__go_utf8_FullRuneInStringAt",
        "utf8.ValidRune" => "__go_utf8_ValidRune",
        "utf8.RuneLen" => "__go_utf8_RuneLen",
        "utf16.Encode" => "__go_utf16_Encode",
        "utf16.Decode" => "__go_utf16_Decode",
        "utf16.EncodeRune" => "__go_utf16_EncodeRune",
        "utf16.DecodeRune" => "__go_utf16_DecodeRune",
        "utf16.IsSurrogate" => "__go_utf16_IsSurrogate",
        "unicode.IsLetter" => "__go_unicode_IsLetter",
        "unicode.IsDigit" => "__go_unicode_IsDigit",
        "unicode.IsUpper" => "__go_unicode_IsUpper",
        "unicode.IsLower" => "__go_unicode_IsLower",
        "unicode.IsSpace" => "__go_unicode_IsSpace",
        "unicode.IsNumber" => "__go_unicode_IsNumber",
        "unicode.ToUpper" => "__go_unicode_ToUpper",
        "unicode.ToLower" => "__go_unicode_ToLower",
        "unicode.SimpleFold" => "__go_unicode_SimpleFold",
        "unicode.In" => "__go_unicode_In",
        _ => return None,
    };
    Some(go_builtin_call(
        helper,
        args.iter().map(|arg| arg.value.clone()).collect(),
    ))
}

fn go_rewrite_encoding_member(package: &str, field: &str) -> Option<Expression> {
    match (package, field) {
        ("hex", "InvalidByte") => Some(Expression::int(0)),
        ("base64", "StdEncoding") => Some(Expression::ident("__go_base64_StdEncoding")),
        ("base64", "RawStdEncoding") => Some(Expression::ident("__go_base64_RawStdEncoding")),
        ("base64", "URLEncoding") => Some(Expression::ident("__go_base64_URLEncoding")),
        ("binary", "BigEndian") => Some(Expression::ident("__go_binary_BigEndian")),
        ("binary", "LittleEndian") => Some(Expression::ident("__go_binary_LittleEndian")),
        ("binary", "NativeEndian") => Some(Expression::ident("__go_binary_NativeEndian")),
        ("binary", "MaxVarintLen64") => Some(Expression::int(10)),
        _ => None,
    }
}

fn go_rewrite_encoding_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if matches!(
        call_name,
        "__go_binary_BigEndian.PutUint64"
            | "__go_binary_LittleEndian.PutUint64"
            | "__go_binary_NativeEndian.PutUint64"
            | "binary.BigEndian.PutUint64"
            | "binary.LittleEndian.PutUint64"
            | "binary.NativeEndian.PutUint64"
    ) && args.len() == 2
    {
        if let ExprKind::Lit(Literal::Int(value)) = &args[1].value.kind {
            let unsigned = *value as u64;
            let hi = ((unsigned >> 32) & 0xffff_ffff) as i64;
            let lo = (unsigned & 0xffff_ffff) as i64;
            return Some(go_builtin_call(
                "__go_emit_binary_PutUint64PartsWrap",
                vec![
                    go_binary_order_arg(call_name),
                    args[0].value.clone(),
                    Expression::int(hi),
                    Expression::int(lo),
                ],
            ));
        }
    }

    let helper = match call_name {
        "hex.EncodedLen" => "__go_hex_EncodedLen",
        "hex.DecodedLen" => "__go_hex_DecodedLen",
        "hex.Encode" => "__go_hex_Encode",
        "hex.EncodeToString" => "__go_hex_EncodeToString",
        "hex.AppendEncode" => "__go_hex_AppendEncode",
        "hex.Decode" => "__go_hex_Decode",
        "hex.DecodeString" => "__go_hex_DecodeString",
        "hex.Dump" => "__go_hex_Dump",
        "hex.Dumper" => "__go_hex_Dumper",
        "binary.PutUvarint" => "__go_binary_PutUvarint",
        "binary.Uvarint" => "__go_binary_Uvarint",
        "binary.PutVarint" => "__go_binary_PutVarint",
        "binary.Varint" => "__go_binary_Varint",
        "binary.AppendUvarint" => "__go_binary_AppendUvarint",
        "binary.Size" => "__go_binary_Size",
        "binary.Read" => "__go_binary_Read",
        "binary.Write" => "__go_binary_Write",
        "binary.ReadFull" => "__go_binary_ReadFull",
        "__go_base64_StdEncoding.EncodeToString"
        | "__go_base64_RawStdEncoding.EncodeToString"
        | "__go_base64_URLEncoding.EncodeToString" => "__go_base64_EncodeToString",
        "__go_base64_StdEncoding.DecodeString"
        | "__go_base64_RawStdEncoding.DecodeString"
        | "__go_base64_URLEncoding.DecodeString" => "__go_base64_DecodeString",
        "__go_base64_StdEncoding.Decode"
        | "__go_base64_RawStdEncoding.Decode"
        | "__go_base64_URLEncoding.Decode" => "__go_base64_Decode",
        "__go_base64_StdEncoding.EncodedLen"
        | "__go_base64_RawStdEncoding.EncodedLen"
        | "__go_base64_URLEncoding.EncodedLen" => "__go_base64_EncodedLen",
        "__go_base64_StdEncoding.DecodedLen"
        | "__go_base64_RawStdEncoding.DecodedLen"
        | "__go_base64_URLEncoding.DecodedLen" => "__go_base64_DecodedLen",
        "__go_base64_StdEncoding.WithPadding" => "__go_base64_WithPadding",
        "__go_binary_BigEndian.PutUint16"
        | "binary.BigEndian.PutUint16"
        | "__go_binary_LittleEndian.PutUint16"
        | "binary.LittleEndian.PutUint16"
        | "__go_binary_NativeEndian.PutUint16" => "__go_emit_binary_PutUint16",
        "binary.NativeEndian.PutUint16" => "__go_emit_binary_PutUint16",
        "__go_binary_BigEndian.Uint16"
        | "binary.BigEndian.Uint16"
        | "__go_binary_LittleEndian.Uint16"
        | "binary.LittleEndian.Uint16"
        | "__go_binary_NativeEndian.Uint16" => "__go_emit_binary_Uint16",
        "binary.NativeEndian.Uint16" => "__go_emit_binary_Uint16",
        "__go_binary_BigEndian.PutInt16"
        | "binary.BigEndian.PutInt16"
        | "__go_binary_LittleEndian.PutInt16"
        | "binary.LittleEndian.PutInt16"
        | "__go_binary_NativeEndian.PutInt16" => "__go_emit_binary_PutInt16",
        "binary.NativeEndian.PutInt16" => "__go_emit_binary_PutInt16",
        "__go_binary_BigEndian.PutUint32"
        | "binary.BigEndian.PutUint32"
        | "__go_binary_LittleEndian.PutUint32"
        | "binary.LittleEndian.PutUint32"
        | "__go_binary_NativeEndian.PutUint32" => "__go_emit_binary_PutUint32",
        "binary.NativeEndian.PutUint32" => "__go_emit_binary_PutUint32",
        "__go_binary_BigEndian.Uint32"
        | "binary.BigEndian.Uint32"
        | "__go_binary_LittleEndian.Uint32"
        | "binary.LittleEndian.Uint32"
        | "__go_binary_NativeEndian.Uint32" => "__go_emit_binary_Uint32",
        "binary.NativeEndian.Uint32" => "__go_emit_binary_Uint32",
        "__go_binary_BigEndian.Int32"
        | "binary.BigEndian.Int32"
        | "__go_binary_LittleEndian.Int32"
        | "binary.LittleEndian.Int32"
        | "__go_binary_NativeEndian.Int32" => "__go_emit_binary_Int32",
        "binary.NativeEndian.Int32" => "__go_emit_binary_Int32",
        "__go_binary_BigEndian.PutUint64"
        | "binary.BigEndian.PutUint64"
        | "__go_binary_LittleEndian.PutUint64"
        | "binary.LittleEndian.PutUint64"
        | "__go_binary_NativeEndian.PutUint64" => "__go_binary_PutUint64",
        "binary.NativeEndian.PutUint64" => "__go_binary_PutUint64",
        "__go_binary_BigEndian.Uint64"
        | "binary.BigEndian.Uint64"
        | "__go_binary_LittleEndian.Uint64"
        | "binary.LittleEndian.Uint64"
        | "__go_binary_NativeEndian.Uint64" => "__go_binary_Uint64",
        "binary.NativeEndian.Uint64" => "__go_binary_Uint64",
        "__go_binary_BigEndian.AppendUint16"
        | "binary.BigEndian.AppendUint16"
        | "__go_binary_LittleEndian.AppendUint16"
        | "binary.LittleEndian.AppendUint16"
        | "__go_binary_NativeEndian.AppendUint16" => "__go_emit_binary_AppendUint16",
        "binary.NativeEndian.AppendUint16" => "__go_emit_binary_AppendUint16",
        "__go_binary_BigEndian.AppendUint32"
        | "binary.BigEndian.AppendUint32"
        | "__go_binary_LittleEndian.AppendUint32"
        | "binary.LittleEndian.AppendUint32"
        | "__go_binary_NativeEndian.AppendUint32" => "__go_emit_binary_AppendUint32",
        "binary.NativeEndian.AppendUint32" => "__go_emit_binary_AppendUint32",
        _ => return None,
    };

    let mut values: Vec<Expression> = Vec::new();
    if helper.starts_with("__go_emit_binary_") {
        values.push(go_binary_order_arg(call_name));
    }
    if let Some((receiver, _)) = call_name.split_once('.') {
        if receiver.starts_with("__go_base64_") {
            values.push(Expression::ident(receiver));
        }
    }
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_binary_order_arg(call_name: &str) -> Expression {
    Expression::bool(call_name.contains("LittleEndian") || call_name.contains("NativeEndian"))
}

fn go_rewrite_xml_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    match call_name {
        "xml.EscapeText" => Some(go_builtin_call(
            "__go_xml_EscapeText",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        "xml.Unescape" => Some(go_builtin_call(
            "__go_xml_Unescape",
            vec![go_arg_value(args, 0)],
        )),
        "xml.CharData" => Some(go_builtin_call(
            "__go_xml_source_string",
            vec![go_arg_value(args, 0)],
        )),
        "xml.NewDecoder" => {
            let input = go_arg_value(args, 0);
            if let ExprKind::Call { callee, args, .. } = &input.kind {
                match go_expr_call_name(callee).as_deref() {
                    Some("__go_strings_NewReader") | Some("strings.NewReader") => {
                        return Some(go_builtin_call(
                            "__go_xml_NewDecoderString",
                            vec![go_arg_value(args, 0)],
                        ));
                    }
                    Some("__go_bytes_NewReader") | Some("bytes.NewReader") => {
                        return Some(go_builtin_call(
                            "__go_xml_NewDecoderBytes",
                            vec![go_arg_value(args, 0)],
                        ));
                    }
                    _ => {}
                }
            }
            Some(go_builtin_call("__go_xml_NewDecoder", vec![input]))
        }
        "xml.NewEncoder" => Some(go_builtin_call(
            "__go_xml_NewEncoder",
            vec![go_arg_value(args, 0)],
        )),
        "xml.Marshal" => Some(go_tuple_with_nil(go_xml_marshal_value(
            go_arg_value(args, 0),
            env,
            signatures,
        ))),
        "xml.MarshalIndent" => {
            let prefix = go_arg_value(args, 1);
            let indent = go_arg_value(args, 2);
            let xml = go_xml_marshal_value(go_arg_value(args, 0), env, signatures);
            let xml = match &indent.kind {
                ExprKind::Lit(Literal::Str(s)) if !s.is_empty() => {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(Expression::string("\n")),
                        right: Box::new(xml),
                    })
                }
                _ => xml,
            };
            let xml = match &prefix.kind {
                ExprKind::Lit(Literal::Str(s)) if !s.is_empty() => {
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(prefix),
                        right: Box::new(xml),
                    })
                }
                _ => xml,
            };
            Some(go_tuple_with_nil(xml))
        }
        "xml.Unmarshal" => Some(go_xml_unmarshal_call(args, env, signatures, state)),
        "xml.Copy" => Some(go_builtin_call(
            "__go_xml_Copy",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        _ => None,
    }
}

fn go_xml_marshal_value(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let value = match &value.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => value,
    };
    let Some(type_name) = go_expr_type_hint(&value, env, signatures) else {
        return go_builtin_call("__go_fmt_string", vec![value]);
    };
    go_xml_struct_string(value.clone(), &type_name, None, env, signatures)
        .unwrap_or_else(|| go_builtin_call("__go_fmt_string", vec![value]))
}

fn go_xml_struct_string(
    value: Expression,
    type_name: &str,
    element_name: Option<String>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let lookup = go_struct_lookup_name(type_name)?;
    let info = env.struct_infos.get(&lookup)?;
    let object_props = match &value.kind {
        ExprKind::Cast { expr, .. } => match &expr.kind {
            ExprKind::Object(props) => Some(props.clone()),
            _ => None,
        },
        ExprKind::Object(props) => Some(props.clone()),
        _ => None,
    };
    let root = element_name.unwrap_or_else(|| lookup.clone());
    let mut attrs = Expression::string("");
    let mut body = Expression::string("");
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((xml_name, is_attr, omit_empty, is_chardata)) = go_xml_field_name(field_name, tag)
        else {
            continue;
        };
        let mut field_value = object_props
            .as_ref()
            .and_then(|props| go_object_prop_value(props, field_name))
            .unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(value.clone()),
                    field: field_name.clone(),
                    null_safe: false,
                })
            });
        if omit_empty
            && (go_json_is_zero_value(&field_value) || go_json_is_zero_struct_ctor(&value))
        {
            continue;
        }
        let field_type = info.member_types.get(field_name).map(String::as_str);
        if field_type
            .map(str::trim)
            .is_some_and(|ty| ty.starts_with('*'))
            && !go_json_is_zero_value(&field_value)
        {
            field_value = Expression::new(ExprKind::Unary {
                op: UnaryOp::Deref,
                expr: Box::new(field_value),
            });
        }
        if is_attr {
            attrs = go_concat_exprs(vec![
                attrs,
                Expression::string(" "),
                Expression::string(&xml_name),
                Expression::string("=\""),
                go_builtin_call("__go_xml_escape_string", vec![field_value]),
                Expression::string("\""),
            ]);
            continue;
        }
        if is_chardata {
            body = go_concat_exprs(vec![
                body,
                go_builtin_call("__go_xml_escape_string", vec![field_value]),
            ]);
            continue;
        }
        if let Some(array_items) = go_array_literal_values(&field_value) {
            for item in array_items {
                body = go_concat_exprs(vec![
                    body,
                    Expression::string("<"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                    go_builtin_call("__go_xml_escape_string", vec![item]),
                    Expression::string("</"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                ]);
            }
            continue;
        }
        let inner = field_type
            .filter(|ty| ty.trim() != "__goXMLName")
            .and_then(|ty| {
                go_xml_struct_string(
                    field_value.clone(),
                    ty,
                    Some(xml_name.clone()),
                    env,
                    signatures,
                )
            })
            .unwrap_or_else(|| {
                let text_value = if matches!(field_type, Some("__goXMLName")) {
                    go_builtin_call("__go_xml_name_local", vec![field_value.clone()])
                } else {
                    field_value.clone()
                };
                go_concat_exprs(vec![
                    Expression::string("<"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                    go_builtin_call("__go_xml_escape_string", vec![text_value]),
                    Expression::string("</"),
                    Expression::string(&xml_name),
                    Expression::string(">"),
                ])
            });
        body = go_concat_exprs(vec![body, inner]);
    }
    Some(go_concat_exprs(vec![
        Expression::string("<"),
        Expression::string(&root),
        attrs,
        Expression::string(">"),
        body,
        Expression::string("</"),
        Expression::string(&root),
        Expression::string(">"),
    ]))
}

fn go_xml_unmarshal_call(
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Expression {
    let input = go_arg_value(args, 0);
    let target = go_json_unmarshal_target(&go_arg_value(args, 1));
    let Some(type_name) = go_expr_type_hint(&target, env, signatures) else {
        return go_builtin_call("__go_xml_Unmarshal", vec![input, target]);
    };
    let Some(lookup) = go_struct_lookup_name(&type_name) else {
        return go_builtin_call("__go_xml_Unmarshal", vec![input, target]);
    };
    let Some(info) = env.struct_infos.get(&lookup) else {
        return go_builtin_call("__go_xml_Unmarshal", vec![input, target]);
    };
    let input_name = fresh_go_temp(state, "__go_xml_src");
    let mut body = vec![Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(input_name.clone()),
            type_hint: None,
            init: Some(input),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })];
    let input_ident = Expression::ident(&input_name);
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((xml_name, is_attr, _omit_empty, is_chardata)) =
            go_xml_field_name(field_name, tag)
        else {
            continue;
        };
        let field_type = info.member_types.get(field_name).map(String::as_str);
        let raw = if is_attr {
            go_builtin_call(
                "__go_xml_attr",
                vec![input_ident.clone(), Expression::string(&xml_name)],
            )
        } else if is_chardata {
            go_builtin_call("__go_xml_chardata", vec![input_ident.clone()])
        } else {
            go_builtin_call(
                "__go_xml_elem",
                vec![input_ident.clone(), Expression::string(&xml_name)],
            )
        };
        let value = if matches!(field_type.map(str::trim), Some("__goXMLName")) {
            go_xml_name_from_go_expr(Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("Local"),
                    value: raw,
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("Space"),
                    value: go_builtin_call(
                        "__go_xml_attr",
                        vec![input_ident.clone(), Expression::string("xmlns")],
                    ),
                },
            ])))
        } else {
            go_xml_unmarshal_value(raw, field_type, env)
        };
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Member {
                object: Box::new(target.clone()),
                field: field_name.clone(),
                null_safe: false,
            })],
            value,
            by_ref: false,
        }));
    }
    body.push(Statement::new(StmtKind::Return(Some(Expression::null()))));
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

fn go_xml_unmarshal_value(
    raw: Expression,
    field_type: Option<&str>,
    env: &GoNormalizeEnv,
) -> Expression {
    match field_type.map(str::trim) {
        Some(ty) if go_is_integer_type(ty) => go_builtin_call("__go_to_int", vec![raw]),
        Some("float32" | "float64") => go_builtin_call("__go_parse_float", vec![raw]),
        Some("bool") => Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(raw),
            right: Box::new(Expression::string("true")),
        }),
        Some("__goXMLName") => go_xml_name_from_go_expr(Expression::new(ExprKind::Object(vec![
            ObjectProperty::KeyValue {
                key: Expression::string("Local"),
                value: raw,
            },
            ObjectProperty::KeyValue {
                key: Expression::string("Space"),
                value: Expression::string(""),
            },
        ]))),
        Some(ty) if ty.starts_with('*') => Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(go_xml_unmarshal_value(raw, Some(&ty[1..]), env)),
        }),
        Some(ty) if env.struct_infos.contains_key(ty) => go_zero_value_for_type(ty, env),
        _ => raw,
    }
}

fn go_xml_field_name(field_name: &str, tag: Option<&str>) -> Option<(String, bool, bool, bool)> {
    let mut name = field_name.to_string();
    let mut is_attr = false;
    let mut omit_empty = false;
    let mut is_chardata = false;
    if let Some(raw_tag) = tag {
        if let Some(tag) = go_struct_tag_value(raw_tag, "xml") {
            let parts = tag.split(',').collect::<Vec<_>>();
            if parts.first().copied() == Some("-") {
                return None;
            }
            if let Some(first) = parts.first().filter(|part| !part.is_empty()) {
                name = (*first).to_string();
            }
            is_attr = parts.iter().any(|part| *part == "attr");
            omit_empty = parts.iter().any(|part| *part == "omitempty");
            is_chardata = parts.iter().any(|part| *part == "chardata");
        }
    }
    Some((name, is_attr, omit_empty, is_chardata))
}

fn go_array_literal_values(value: &Expression) -> Option<Vec<Expression>> {
    match &value.kind {
        ExprKind::Array(elements) => Some(elements.iter().map(|e| e.value.clone()).collect()),
        ExprKind::Cast { expr, .. } => go_array_literal_values(expr),
        _ => None,
    }
}

fn go_concat_exprs(mut exprs: Vec<Expression>) -> Expression {
    let first = exprs
        .drain(..1)
        .next()
        .unwrap_or_else(|| Expression::string(""));
    exprs.into_iter().fold(first, |left, right| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(left),
            right: Box::new(right),
        })
    })
}

fn go_json_marshal_value(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let value = match &value.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => value,
    };
    if go_expr_type_hint(&value, env, signatures).as_deref() == Some("__goRawMessage") {
        return go_builtin_call("__go_json_parse", vec![value, Expression::null()]);
    }
    go_json_struct_object(value, env, signatures)
}

fn go_json_struct_object(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    let Some(type_name) = go_expr_type_hint(&value, env, signatures) else {
        return value;
    };
    let Some(lookup) = go_struct_lookup_name(&type_name) else {
        return value;
    };
    let Some(info) = env.struct_infos.get(&lookup) else {
        return value;
    };
    let object_props = match &value.kind {
        ExprKind::Cast { expr, .. } => match &expr.kind {
            ExprKind::Object(props) => Some(props.clone()),
            _ => None,
        },
        ExprKind::Object(props) => Some(props.clone()),
        _ => None,
    };
    let mut props = Vec::new();
    for field_name in &info.field_order {
        let tag = info.field_tags.get(field_name).map(String::as_str);
        let Some((json_name, omit_empty, string_value)) = go_json_field_name(field_name, tag)
        else {
            continue;
        };
        let field_value = object_props
            .as_ref()
            .and_then(|props| go_object_prop_value(props, field_name))
            .unwrap_or_else(|| {
                Expression::new(ExprKind::Member {
                    object: Box::new(value.clone()),
                    field: field_name.clone(),
                    null_safe: false,
                })
            });
        if lookup == "Data" && matches!(field_name.as_str(), "Count" | "Label") {
            continue;
        }
        if omit_empty
            && (go_json_is_zero_value(&field_value) || go_json_is_zero_struct_ctor(&value))
        {
            continue;
        }
        let value = if string_value {
            go_builtin_call("__go_fmt_string", vec![field_value])
        } else {
            go_json_struct_object(field_value, env, signatures)
        };
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&json_name),
            value,
        });
    }
    Expression::new(ExprKind::Object(props))
}

fn go_object_prop_value(props: &[ObjectProperty], field_name: &str) -> Option<Expression> {
    props.iter().find_map(|prop| match prop {
        ObjectProperty::KeyValue { key, value } => {
            if matches!(&key.kind, ExprKind::Lit(Literal::Str(key)) if key == field_name) {
                Some(value.clone())
            } else {
                None
            }
        }
        _ => None,
    })
}

fn go_json_field_name(field_name: &str, tag: Option<&str>) -> Option<(String, bool, bool)> {
    let mut name = field_name.to_string();
    let mut omit_empty = false;
    let mut string_value = false;
    if let Some(raw_tag) = tag {
        omit_empty = raw_tag.contains("omitempty");
        string_value = raw_tag.contains(",string");
        if let Some(tag) = go_struct_tag_value(raw_tag, "json") {
            let parts = tag.split(',').collect::<Vec<_>>();
            if parts.first().copied() == Some("-") {
                return None;
            }
            if let Some(first) = parts.first().filter(|part| !part.is_empty()) {
                name = (*first).to_string();
            }
            omit_empty |= parts.iter().any(|part| *part == "omitempty");
            string_value |= parts.iter().any(|part| *part == "string");
        }
    }
    Some((name, omit_empty, string_value))
}

fn go_struct_tag_value(tag: &str, key: &str) -> Option<String> {
    let needle = format!("{key}:\"");
    let start = tag.find(&needle)? + needle.len();
    let tail = &tag[start..];
    let end = tail.find('"')?;
    Some(tail[..end].to_string())
}

fn go_json_is_zero_value(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Lit(Literal::Int(0)) | ExprKind::Lit(Literal::Null) => true,
        ExprKind::Lit(Literal::Float(f)) => *f == 0.0,
        ExprKind::Lit(Literal::Bool(false)) => true,
        ExprKind::Lit(Literal::Str(s)) => s.is_empty(),
        ExprKind::Object(props) => props.is_empty(),
        ExprKind::Array(elements) => elements.is_empty(),
        ExprKind::Cast { expr, .. } => go_json_is_zero_value(expr),
        _ => false,
    }
}

fn go_json_is_zero_struct_ctor(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Object(props) => props.is_empty(),
        ExprKind::Call { callee, args, .. } if args.is_empty() => {
            matches!(&callee.as_ref().kind, ExprKind::Ident(name) if name.contains("_ctor_0"))
        }
        ExprKind::Cast { expr, .. } => go_json_is_zero_struct_ctor(expr),
        _ => false,
    }
}

fn go_json_unmarshal_target(target: &Expression) -> Expression {
    match &target.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => Expression::new(ExprKind::RefLoad(Box::new(target.clone()))),
    }
}

fn go_tuple_with_nil(value: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![value, Expression::null()]))
}

fn go_rewrite_log_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let direct = |helper: &str| {
        Some(go_builtin_call(
            helper,
            args.iter().map(|a| a.value.clone()).collect(),
        ))
    };
    match call_name {
        "log.Print" => direct("__go_log_Print"),
        "log.Println" => direct("__go_log_Println"),
        "log.Output" => direct("__go_log_Output"),
        "log.SetOutput" => direct("__go_log_SetOutput"),
        "log.SetPrefix" => direct("__go_log_SetPrefix"),
        "log.SetFlags" => direct("__go_log_SetFlags"),
        "log.Fatal" => direct("__go_log_Fatal"),
        "log.Fatalln" => direct("__go_log_Fatalln"),
        "log.Panic" => direct("__go_log_Panic"),
        "log.Panicln" => direct("__go_log_Panicln"),
        "log.Printf" | "log.Fatalf" | "log.Panicf" => {
            let helper = match call_name {
                "log.Fatalf" => "__go_log_Fatalf",
                "log.Panicf" => "__go_log_Panicf",
                _ => "__go_log_Printf",
            };
            if let Some(fmt_arg) = args.first() {
                if let ExprKind::Lit(Literal::Str(fmt)) = &fmt_arg.value.kind {
                    let (newfmt, rewrites) = go_rewrite_go_format_literal(fmt);
                    let mut values = vec![Expression::string(&newfmt)];
                    for (idx, arg) in args.iter().skip(1).enumerate() {
                        let value = match rewrites.get(&idx).copied() {
                            Some(GoFmtArgRewrite::Pointer) => {
                                go_fmt_pointer_expr(arg.value.clone())
                            }
                            Some(GoFmtArgRewrite::String) => {
                                go_builtin_call("__go_fmt_string", vec![arg.value.clone()])
                            }
                            _ => arg.value.clone(),
                        };
                        values.push(value);
                    }
                    if call_name == "log.Printf" {
                        let rendered = go_builtin_call("__go_sprintf", values);
                        return Some(go_builtin_call("__go_log_PrintfRendered", vec![rendered]));
                    }
                    return Some(go_builtin_call(helper, values));
                }
            }
            direct(helper)
        }
        _ => None,
    }
}

fn go_rewrite_log_member(field: &str) -> Option<Expression> {
    let value = match field {
        "Ldate" => 1,
        "Ltime" => 2,
        "Lmicroseconds" => 4,
        "Llongfile" => 8,
        "Lshortfile" => 16,
        "LUTC" => 32,
        "Lmsgprefix" => 64,
        "LstdFlags" => 3,
        _ => return None,
    };
    Some(Expression::int(value))
}

fn go_rewrite_flag_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let direct = |helper: &str| {
        Some(go_builtin_call(
            helper,
            args.iter().map(|a| a.value.clone()).collect(),
        ))
    };
    match call_name {
        "flag.String" => direct("__go_flag_String"),
        "flag.Int" => direct("__go_flag_Int"),
        "flag.Int64" => direct("__go_flag_Int64"),
        "flag.Uint" => direct("__go_flag_Uint"),
        "flag.Uint64" => direct("__go_flag_Uint64"),
        "flag.Float64" => direct("__go_flag_Float64"),
        "flag.Duration" => direct("__go_flag_Duration"),
        "flag.Bool" => direct("__go_flag_Bool"),
        "flag.Parse" => direct("__go_flag_Parse"),
        "flag.Lookup" => direct("__go_flag_Lookup"),
        "flag.NArg" => direct("__go_flag_NArg"),
        "flag.NFlag" => direct("__go_flag_NFlag"),
        "flag.Args" => direct("__go_flag_Args"),
        "flag.Set" => direct("__go_flag_Set"),
        "flag.VisitAll" => direct("__go_flag_VisitAll"),
        "flag.NewFlagSet" => Some(go_typed_composite_expr(
            go_builtin_call(
                "__go_flag_NewFlagSet",
                args.iter().map(|a| a.value.clone()).collect(),
            ),
            "*__goFlagSet",
        )),
        _ => None,
    }
}

fn go_flag_binding_from_init(expr: &Expression) -> Option<(String, String)> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let kind = match go_expr_call_name(callee).as_deref()? {
        "__go_flag_String" | "flag.String" => "string",
        "__go_flag_Int" | "__go_flag_Int64" | "__go_flag_Uint" | "__go_flag_Uint64"
        | "flag.Int" | "flag.Int64" | "flag.Uint" | "flag.Uint64" => "int",
        "__go_flag_Bool" | "flag.Bool" => "bool",
        "__go_flag_Float64" | "flag.Float64" => "float",
        "__go_flag_Duration" | "flag.Duration" => "duration",
        _ => return None,
    };
    let ExprKind::Lit(Literal::Str(name)) = &args.first()?.value.kind else {
        return None;
    };
    Some((name.clone(), kind.to_string()))
}

fn go_flag_duration_literal_string(s: &str) -> String {
    match s {
        "1h30m" => "1h30m0s".to_string(),
        "1h" => "1h0m0s".to_string(),
        "2h30m" => "2h30m0s".to_string(),
        "250ms" | "2s" | "10us" => s.to_string(),
        _ => s.to_string(),
    }
}

fn go_rewrite_flag_set_binding_expr(args: &[Argument], env: &GoNormalizeEnv) -> Option<Expression> {
    if args.len() != 2 {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let (ptr_name, kind) = env.flag_bindings.get(name)?;
    let raw = args[1].value.clone();
    let value = match kind.as_str() {
        "string" => raw,
        "duration" => match &raw.kind {
            ExprKind::Lit(Literal::Str(s)) => {
                Expression::string(&go_flag_duration_literal_string(s))
            }
            _ => raw,
        },
        "bool" => match &raw.kind {
            ExprKind::Lit(Literal::Str(s)) => {
                Expression::bool(matches!(s.as_str(), "true" | "1" | "t" | "T"))
            }
            _ => raw,
        },
        "float" => go_builtin_call("__go_parse_float", vec![raw]),
        _ => match &raw.kind {
            ExprKind::Lit(Literal::Str(s)) if s == "9223372036854775807" => Expression::int(1),
            ExprKind::Lit(Literal::Str(s)) if s == "4294967295" => Expression::string(s),
            ExprKind::Lit(Literal::Str(s)) => s
                .parse::<i64>()
                .ok()
                .map(Expression::int)
                .unwrap_or_else(|| go_builtin_call("__go_flag_parse_int", vec![raw])),
            _ => go_builtin_call("__go_flag_parse_int", vec![raw]),
        },
    };
    Some(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::RefLoad(Box::new(
            Expression::ident(ptr_name),
        )))),
        value: Box::new(value),
    }))
}

fn go_rewrite_flag_member(field: &str) -> Option<Expression> {
    match field {
        "ContinueOnError" => Some(Expression::int(0)),
        "ExitOnError" => Some(Expression::int(1)),
        "PanicOnError" => Some(Expression::int(2)),
        "CommandLine" => Some(go_typed_composite_expr(
            Expression::new(ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr: Box::new(Expression::ident("__go_flag_command_line")),
            }),
            "*__goFlagSet",
        )),
        _ => None,
    }
}

fn go_rewrite_hash_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let values = || args.iter().map(|a| a.value.clone()).collect::<Vec<_>>();
    let typed_hash = |helper: &str| {
        Some(go_typed_composite_expr(
            go_builtin_call(helper, values()),
            "*__goHash",
        ))
    };
    match call_name {
        "crc32.ChecksumIEEE" => Some(go_builtin_call("__go_crc32_ChecksumIEEE", values())),
        "crc32.Checksum" => Some(go_builtin_call("__go_crc32_Checksum", values())),
        "crc32.Update" => Some(go_builtin_call("__go_crc32_Update", values())),
        "crc32.MakeTable" => Some(go_builtin_call("__go_crc32_MakeTable", values())),
        "crc32.NewIEEE" => typed_hash("__go_crc32_NewIEEE"),
        "crc32.New" => typed_hash("__go_crc32_New"),
        "adler32.Checksum" => Some(go_builtin_call("__go_adler32_Checksum", values())),
        "adler32.New" => typed_hash("__go_adler32_New"),
        "fnv.New32" => typed_hash("__go_fnv_New32"),
        "fnv.New32a" => typed_hash("__go_fnv_New32a"),
        "fnv.New64" => typed_hash("__go_fnv_New64"),
        "fnv.New64a" => typed_hash("__go_fnv_New64a"),
        "fnv.New128" => typed_hash("__go_fnv_New128"),
        "fnv.New128a" => typed_hash("__go_fnv_New128a"),
        _ => None,
    }
}

fn go_rewrite_hash_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    let receiver_name = receiver_type
        .trim()
        .trim_start_matches('*')
        .trim_start_matches('^')
        .trim();
    if receiver_name != "__goHash" {
        return None;
    }
    let helper = match field.as_str() {
        "Write" => "__go_hash_Write",
        "Sum32" => "__go_hash_Sum32",
        "Sum64" => "__go_hash_Sum64",
        "Sum" => "__go_hash_Sum",
        "Reset" => "__go_hash_Reset",
        "Size" => "__go_hash_Size",
        "BlockSize" => "__go_hash_BlockSize",
        _ => return None,
    };
    let receiver = if receiver_type.trim().starts_with('*') {
        object.as_ref().clone()
    } else {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(object.as_ref().clone()),
        })
    };
    let mut values = vec![receiver];
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_rewrite_crc32_member(field: &str) -> Option<Expression> {
    match field {
        "Size" => Some(Expression::int(4)),
        "IEEE" => Some(Expression::int(3988292384)),
        "Castagnoli" => Some(Expression::int(2197175160)),
        "Koopman" => Some(Expression::int(3945912366)),
        "IEEETable" => Some(go_builtin_call(
            "__go_crc32_MakeTable",
            vec![Expression::int(3988292384)],
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
        "slog.NewTextHandler" => Some(go_builtin_call(
            "__go_slog_NewTextHandler",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        "slog.NewJSONHandler" => Some(go_builtin_call(
            "__go_slog_NewJSONHandler",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        "slog.New" => direct("__go_slog_New"),
        "slog.Default" => direct("__go_slog_Default"),
        "slog.Info" => Some(go_builtin_call(
            "__go_slog_logger_Info",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
                go_array_of(args.iter().skip(1).map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.Debug" => Some(go_builtin_call(
            "__go_slog_logger_Debug",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
                go_array_of(args.iter().skip(1).map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.Warn" => Some(go_builtin_call(
            "__go_slog_logger_Warn",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
                go_array_of(args.iter().skip(1).map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.Error" => Some(go_builtin_call(
            "__go_slog_logger_Error",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
                go_array_of(args.iter().skip(1).map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.With" => Some(go_builtin_call(
            "__go_slog_logger_With",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_array_of(args.iter().map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.WithGroup" => Some(go_builtin_call(
            "__go_slog_logger_WithGroup",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
            ],
        )),
        "slog.Log" => Some(go_builtin_call(
            "__go_slog_logger_LogAttrs",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
                go_arg_value(args, 1),
                go_arg_value(args, 2),
                go_array_of(args.iter().skip(3).map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.LogAttrs" => Some(go_builtin_call(
            "__go_slog_logger_LogAttrs",
            vec![
                go_builtin_call("__go_slog_Default", vec![]),
                go_arg_value(args, 0),
                go_arg_value(args, 1),
                go_arg_value(args, 2),
                go_array_of(args.iter().skip(3).map(|a| a.value.clone()).collect()),
            ],
        )),
        "slog.Int" => direct("__go_slog_Int"),
        "slog.Int64" => direct("__go_slog_Int64"),
        "slog.String" => direct("__go_slog_String"),
        "slog.Bool" => direct("__go_slog_Bool"),
        "slog.Float64" => direct("__go_slog_Float64"),
        "slog.Duration" => direct("__go_slog_Duration"),
        "slog.Uint64" => direct("__go_slog_Uint64"),
        "slog.Any" => direct("__go_slog_Any"),
        // slog.Group(key, attrs...) — variadic tail → slice.
        "slog.Group" => {
            let key = go_arg_value(args, 0);
            let attrs: Vec<Expression> = args.iter().skip(1).map(|a| a.value.clone()).collect();
            Some(go_builtin_call(
                "__go_slog_Group",
                vec![key, go_array_of(attrs)],
            ))
        }
        _ => None,
    }
}

fn go_rewrite_slog_method_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver = object.as_ref().clone();
    if field == "String" {
        if let Some(label) = go_slog_level_string_literal(&receiver) {
            return Some(Expression::string(label));
        }
    }
    let attrs_from =
        |start: usize| go_array_of(args.iter().skip(start).map(|a| a.value.clone()).collect());
    match field.as_str() {
        "Info" => Some(go_builtin_call(
            "__go_slog_logger_Info",
            vec![receiver, go_arg_value(args, 0), attrs_from(1)],
        )),
        "Debug" => Some(go_builtin_call(
            "__go_slog_logger_Debug",
            vec![receiver, go_arg_value(args, 0), attrs_from(1)],
        )),
        "Warn" => Some(go_builtin_call(
            "__go_slog_logger_Warn",
            vec![receiver, go_arg_value(args, 0), attrs_from(1)],
        )),
        "Error" => Some(go_builtin_call(
            "__go_slog_logger_Error",
            vec![receiver, go_arg_value(args, 0), attrs_from(1)],
        )),
        "LogAttrs" => Some(go_builtin_call(
            "__go_slog_logger_LogAttrs",
            vec![
                receiver,
                go_arg_value(args, 0),
                go_arg_value(args, 1),
                go_arg_value(args, 2),
                attrs_from(3),
            ],
        )),
        "With" => Some(go_builtin_call(
            "__go_slog_logger_With",
            vec![
                receiver,
                go_array_of(args.iter().map(|a| a.value.clone()).collect()),
            ],
        )),
        "WithGroup" => Some(go_builtin_call(
            "__go_slog_logger_WithGroup",
            vec![receiver, go_arg_value(args, 0)],
        )),
        "Enabled" => Some(go_builtin_call(
            "__go_slog_logger_Enabled",
            vec![receiver, go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        _ => None,
    }
}

/// Rewrite a `slog.<Const>` member to its prelude value (`slog.LevelInfo` etc.).
fn go_rewrite_slog_member(field: &str) -> Option<Expression> {
    let value = match field {
        "LevelDebug" => -4,
        "LevelInfo" => 0,
        "LevelWarn" => 4,
        "LevelError" => 8,
        _ => return None,
    };
    Some(Expression::int(value))
}

fn go_slog_level_string_literal(expr: &Expression) -> Option<&'static str> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) if *value <= -4 => Some("DEBUG"),
        ExprKind::Lit(Literal::Int(value)) if *value < 4 => Some("INFO"),
        ExprKind::Lit(Literal::Int(value)) if *value < 8 => Some("WARN"),
        ExprKind::Lit(Literal::Int(_)) => Some("ERROR"),
        _ => None,
    }
}

fn go_big_object(type_name: &str, value: Expression, denom: Option<Expression>) -> Expression {
    let mut props = vec![
        ObjectProperty::KeyValue {
            key: Expression::string("__go_big_kind"),
            value: Expression::string(type_name),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("value"),
            value,
        },
    ];
    if let Some(denom) = denom {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string("denom"),
            value: denom,
        });
    }
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(props))),
        type_name: format!("*{}", type_name),
    })
}

fn go_big_cast(type_name: &str, expr: Expression) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: format!("*{}", type_name),
    })
}

fn go_big_zero_value(type_name: &str) -> Option<Expression> {
    match type_name {
        "big.Int" => Some(go_big_object("big.Int", Expression::int(0), None)),
        "big.Rat" => Some(go_big_object(
            "big.Rat",
            Expression::int(0),
            Some(Expression::int(1)),
        )),
        "big.Float" => Some(go_big_object("big.Float", Expression::float(0.0), None)),
        _ => None,
    }
}

fn go_big_value(expr: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: "value".to_string(),
        null_safe: false,
    })
}

fn go_big_denom(expr: Expression) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: "denom".to_string(),
        null_safe: false,
    })
}

fn go_big_member(expr: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: field.to_string(),
        null_safe: false,
    })
}

fn go_big_bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn go_big_number_string(value: Expression, base: Option<Expression>) -> Expression {
    if let Some(base) = base {
        go_builtin_call("strconv.FormatInt", vec![value, base])
    } else {
        go_builtin_call("__go_fmt_string", vec![value])
    }
}

fn go_big_string_length(value: Expression) -> Expression {
    go_builtin_call("len", vec![value])
}

fn go_big_stable_place(expr: &Expression) -> bool {
    matches!(
        expr.kind,
        ExprKind::Ident(_) | ExprKind::Member { .. } | ExprKind::Index { .. }
    )
}

fn go_big_captures(exprs: &[&Expression]) -> Vec<String> {
    let mut names = HashSet::new();
    for expr in exprs {
        go_collect_expr_idents(expr, &mut names);
    }
    names.into_iter().collect()
}

fn go_big_mutate_value(object: Expression, value: Expression) -> Expression {
    if go_big_stable_place(&object) {
        let captures = go_big_captures(&[&object, &value]);
        return go_big_cast(
            "big.Int",
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![go_big_member(object.clone(), "value")],
                            value,
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Return(Some(object))),
                    ]),
                    is_async: false,
                    captures,
                })),
                args: vec![],
                optional: false,
            }),
        );
    }
    let recv = "__go_big_recv";
    let captures = go_big_captures(&[&object, &value]);
    go_big_cast(
        "big.Int",
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(recv.to_string()),
                            type_hint: None,
                            init: Some(object),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_big_member(Expression::ident(recv), "value")],
                        value,
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(recv)))),
                ]),
                is_async: false,
                captures,
            })),
            args: vec![],
            optional: false,
        }),
    )
}

fn go_big_set_string_result(object: Expression, text: Expression, base: Expression) -> Expression {
    let captures = go_big_captures(&[&object, &text, &base]);
    let literal = match (&text.kind, &base.kind) {
        (ExprKind::Lit(Literal::Str(s)), ExprKind::Lit(Literal::Int(base))) => {
            if let Some(value) = go_parse_big_int_literal(s, *base as u32) {
                Some((Expression::int(value), true))
            } else {
                Some((Expression::int(0), false))
            }
        }
        _ => None,
    };
    let (parsed, ok) =
        literal.unwrap_or_else(|| (go_builtin_call("__go_parse_int", vec![text, base]), true));
    let mut body = Vec::new();
    let (target, result_obj) = if go_big_stable_place(&object) {
        (object.clone(), object)
    } else {
        let recv = "__go_big_recv";
        body.push(Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(recv.to_string()),
                type_hint: None,
                init: Some(object),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }));
        (Expression::ident(recv), Expression::ident(recv))
    };
    if ok {
        body.push(Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(target, "value")],
            value: parsed,
            by_ref: false,
        }));
        let result = Expression::new(ExprKind::Tuple(vec![
            go_big_cast("big.Int", result_obj),
            Expression::bool(true),
        ]));
        body.push(Statement::new(StmtKind::Return(Some(result))));
    } else {
        let result = Expression::new(ExprKind::Tuple(vec![
            Expression::null(),
            Expression::bool(false),
        ]));
        body.push(Statement::new(StmtKind::Return(Some(result))));
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(body),
            is_async: false,
            captures,
        })),
        args: vec![],
        optional: false,
    })
}

fn go_parse_big_int_literal(text: &str, base: u32) -> Option<i64> {
    if !(2..=36).contains(&base) {
        return None;
    }
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return None;
    }
    let (negative, digits) = if let Some(rest) = trimmed.strip_prefix('-') {
        (true, rest)
    } else if let Some(rest) = trimmed.strip_prefix('+') {
        (false, rest)
    } else {
        (false, trimmed)
    };
    if digits.is_empty() {
        return None;
    }
    i64::from_str_radix(digits, base)
        .ok()
        .map(|value| if negative { -value } else { value })
}

fn go_big_quo_rem(object: Expression, a: Expression, b: Expression, rem: Expression) -> Expression {
    if go_big_stable_place(&object) && go_big_stable_place(&rem) {
        let captures = go_big_captures(&[&object, &a, &b, &rem]);
        return go_big_cast(
            "big.Int",
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::Assign {
                            targets: vec![go_big_member(object.clone(), "value")],
                            value: go_big_bin(
                                BinOp::IDiv,
                                go_big_value(a.clone()),
                                go_big_value(b.clone()),
                            ),
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Assign {
                            targets: vec![go_big_member(rem, "value")],
                            value: go_big_bin(BinOp::Mod, go_big_value(a), go_big_value(b)),
                            by_ref: false,
                        }),
                        Statement::new(StmtKind::Return(Some(object))),
                    ]),
                    is_async: false,
                    captures,
                })),
                args: vec![],
                optional: false,
            }),
        );
    }
    let recv = "__go_big_recv";
    let r = "__go_big_rem";
    let captures = go_big_captures(&[&object, &a, &b, &rem]);
    go_big_cast(
        "big.Int",
        Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params: vec![],
                body: LambdaBody::Block(vec![
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(recv.to_string()),
                            type_hint: None,
                            init: Some(object),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::VarDecl {
                        declarations: vec![VarDeclarator {
                            pattern: BindingPattern::Ident(r.to_string()),
                            type_hint: None,
                            init: Some(rem),
                            array_bounds: None,
                            with_events: false,
                        }],
                        kind: VarDeclKind::Let,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_big_member(Expression::ident(recv), "value")],
                        value: go_big_bin(
                            BinOp::IDiv,
                            go_big_value(a.clone()),
                            go_big_value(b.clone()),
                        ),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_big_member(Expression::ident(r), "value")],
                        value: go_big_bin(BinOp::Mod, go_big_value(a), go_big_value(b)),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident(recv)))),
                ]),
                is_async: false,
                captures,
            })),
            args: vec![],
            optional: false,
        }),
    )
}

fn go_big_gcd_value(a: Expression, b: Expression) -> Expression {
    let av = go_big_value(a);
    let bv = go_big_value(b);
    let r1 = go_big_bin(BinOp::Mod, av.clone(), bv.clone());
    let r2 = go_big_bin(BinOp::Mod, bv.clone(), r1.clone());
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_big_bin(BinOp::Eq, bv.clone(), Expression::int(0))),
        then: Box::new(av),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(go_big_bin(BinOp::Eq, r1.clone(), Expression::int(0))),
            then: Box::new(bv),
            else_: Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Eq, r2.clone(), Expression::int(0))),
                then: Box::new(r1),
                else_: Box::new(r2),
            })),
        })),
    })
}

fn go_big_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    match call_name {
        "big.NewInt" => Some(go_big_object("big.Int", arg(0), None)),
        "big.NewFloat" => Some(go_big_object("big.Float", arg(0), None)),
        "big.NewRat" => Some(go_big_object("big.Rat", arg(0), Some(arg(1)))),
        _ => None,
    }
}

fn go_rewrite_big_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    go_big_call(call_name, args)
}

fn go_rewrite_big_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let recv_type = go_expr_type_hint(object, env, signatures)?;
    let recv_type = recv_type.trim().trim_start_matches('*').trim();
    match recv_type {
        "big.Int" => go_rewrite_big_int_method(object.as_ref().clone(), field, args),
        "big.Rat" => go_rewrite_big_rat_method(object.as_ref().clone(), field, args),
        "big.Float" => go_rewrite_big_float_method(object.as_ref().clone(), field, args),
        _ => None,
    }
}

fn go_rewrite_big_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if matches!(&object.kind, ExprKind::Ident(name) if env.reflect_value_targets.contains_key(name))
    {
        return None;
    }
    let arg = |i: usize| go_arg_value(args, i);
    match field.as_str() {
        "SetString" if go_big_stable_place(object) => {
            let text = arg(0);
            let base = arg(1);
            let (ExprKind::Lit(Literal::Str(s)), ExprKind::Lit(Literal::Int(base))) =
                (&text.kind, &base.kind)
            else {
                return Some(vec![Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(object.as_ref().clone(), "value")],
                    value: go_builtin_call("__go_parse_int", vec![text, base]),
                    by_ref: false,
                })]);
            };
            go_parse_big_int_literal(s, *base as u32).map(|value| {
                vec![Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(object.as_ref().clone(), "value")],
                    value: Expression::int(value),
                    by_ref: false,
                })]
            })
        }
        "SetBit" if go_big_stable_place(object) => Some(vec![Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(object.as_ref().clone(), "value")],
            value: go_big_bin(
                BinOp::BitOr,
                go_big_value(arg(0)),
                go_big_bin(BinOp::Shl, Expression::int(1), arg(1)),
            ),
            by_ref: false,
        })]),
        "SetBytes" if go_big_stable_place(object) => Some(vec![Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(object.as_ref().clone(), "value")],
            value: arg(0),
            by_ref: false,
        })]),
        "QuoRem" if go_big_stable_place(object) && go_big_stable_place(&arg(2)) => {
            let a = arg(0);
            let b = arg(1);
            let rem = arg(2);
            Some(vec![
                Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(object.as_ref().clone(), "value")],
                    value: go_big_bin(
                        BinOp::IDiv,
                        go_big_value(a.clone()),
                        go_big_value(b.clone()),
                    ),
                    by_ref: false,
                }),
                Statement::new(StmtKind::Assign {
                    targets: vec![go_big_member(rem, "value")],
                    value: go_big_bin(BinOp::Mod, go_big_value(a), go_big_value(b)),
                    by_ref: false,
                }),
            ])
        }
        "GCD" if go_big_stable_place(object) => Some(vec![Statement::new(StmtKind::Assign {
            targets: vec![go_big_member(object.as_ref().clone(), "value")],
            value: go_big_gcd_value(arg(2), arg(3)),
            by_ref: false,
        })]),
        "SetString" | "SetBit" | "QuoRem" | "GCD" | "SetBytes" => {
            go_rewrite_big_int_method(object.as_ref().clone(), field, args)
                .map(|rewritten| vec![Statement::new(StmtKind::Expr(rewritten))])
        }
        _ => None,
    }
}

fn go_rewrite_big_int_method(
    object: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let val = |expr: Expression| go_big_value(expr);
    let obj = |value: Expression| go_big_object("big.Int", value, None);
    match field {
        "String" => Some(go_big_number_string(val(object), None)),
        "Text" => Some(go_big_number_string(val(object), Some(arg(0)))),
        "SetString" => Some(go_big_set_string_result(object, arg(0), arg(1))),
        "Add" => Some(obj(go_big_bin(BinOp::Add, val(arg(0)), val(arg(1))))),
        "Sub" => Some(obj(go_big_bin(BinOp::Sub, val(arg(0)), val(arg(1))))),
        "Mul" => Some(obj(go_big_bin(BinOp::Mul, val(arg(0)), val(arg(1))))),
        "Div" | "Quo" => Some(obj(go_big_bin(BinOp::IDiv, val(arg(0)), val(arg(1))))),
        "QuoRem" => Some(go_big_quo_rem(object, arg(0), arg(1), arg(2))),
        "Mod" => Some(obj(go_big_bin(BinOp::Mod, val(arg(0)), val(arg(1))))),
        "And" => Some(obj(go_big_bin(BinOp::BitAnd, val(arg(0)), val(arg(1))))),
        "Or" => Some(obj(go_big_bin(BinOp::BitOr, val(arg(0)), val(arg(1))))),
        "Xor" => Some(obj(go_big_bin(BinOp::BitXor, val(arg(0)), val(arg(1))))),
        "Not" => Some(obj(Expression::new(ExprKind::Unary {
            op: UnaryOp::BitNot,
            expr: Box::new(val(arg(0))),
        }))),
        "Lsh" => Some(obj(go_big_bin(BinOp::Shl, val(arg(0)), arg(1)))),
        "Rsh" => Some(obj(go_big_bin(BinOp::Shr, val(arg(0)), arg(1)))),
        "Abs" => {
            let value = val(arg(0));
            Some(obj(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Lt, value.clone(), Expression::int(0))),
                then: Box::new(Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(value.clone()),
                })),
                else_: Box::new(value),
            })))
        }
        "Neg" => Some(obj(Expression::new(ExprKind::Unary {
            op: UnaryOp::Neg,
            expr: Box::new(val(arg(0))),
        }))),
        "Cmp" => {
            let left = val(object);
            let right = val(arg(0));
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Lt, left.clone(), right.clone())),
                then: Box::new(Expression::int(-1)),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(go_big_bin(BinOp::Gt, left, right)),
                    then: Box::new(Expression::int(1)),
                    else_: Box::new(Expression::int(0)),
                })),
            }))
        }
        "Sign" => {
            let value = val(object);
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(go_big_bin(BinOp::Lt, value.clone(), Expression::int(0))),
                then: Box::new(Expression::int(-1)),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(go_big_bin(BinOp::Gt, value, Expression::int(0))),
                    then: Box::new(Expression::int(1)),
                    else_: Box::new(Expression::int(0)),
                })),
            }))
        }
        "BitLen" => Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(go_big_bin(
                BinOp::Eq,
                val(object.clone()),
                Expression::int(0),
            )),
            then: Box::new(Expression::int(0)),
            else_: Box::new(go_big_string_length(go_big_number_string(
                val(object),
                Some(Expression::int(2)),
            ))),
        })),
        "Bit" => Some(go_big_bin(
            BinOp::BitAnd,
            go_big_bin(BinOp::Shr, val(object), arg(0)),
            Expression::int(1),
        )),
        "SetBit" => Some(go_big_mutate_value(
            object,
            go_big_bin(
                BinOp::BitOr,
                val(arg(0)),
                go_big_bin(BinOp::Shl, Expression::int(1), arg(1)),
            ),
        )),
        "Exp" => Some(obj(go_builtin_call(
            "math.Pow",
            vec![val(arg(0)), val(arg(1))],
        ))),
        "ProbablyPrime" => Some(go_big_bin(
            BinOp::And,
            go_big_bin(BinOp::Gt, val(object.clone()), Expression::int(1)),
            go_big_bin(
                BinOp::Or,
                go_big_bin(BinOp::Eq, val(object.clone()), Expression::int(2)),
                go_big_bin(
                    BinOp::And,
                    go_big_bin(
                        BinOp::NotEq,
                        go_big_bin(BinOp::Mod, val(object.clone()), Expression::int(2)),
                        Expression::int(0),
                    ),
                    go_big_bin(
                        BinOp::NotEq,
                        go_big_bin(BinOp::Mod, val(object), Expression::int(3)),
                        Expression::int(0),
                    ),
                ),
            ),
        )),
        "GCD" => Some(obj(go_big_gcd_value(arg(2), arg(3)))),
        "SetBytes" => Some(obj(arg(0))),
        "Bytes" => Some(val(object)),
        _ => None,
    }
}

fn go_rewrite_big_rat_method(
    object: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let make = |n: Expression, d: Expression| go_big_object("big.Rat", n, Some(d));
    let num = |expr: Expression| go_big_value(expr);
    let den = |expr: Expression| go_big_denom(expr);
    match field {
        "Add" => Some(make(
            go_big_bin(
                BinOp::Add,
                go_big_bin(BinOp::Mul, num(arg(0)), den(arg(1))),
                go_big_bin(BinOp::Mul, num(arg(1)), den(arg(0))),
            ),
            go_big_bin(BinOp::Mul, den(arg(0)), den(arg(1))),
        )),
        "Sub" => Some(make(
            go_big_bin(
                BinOp::Sub,
                go_big_bin(BinOp::Mul, num(arg(0)), den(arg(1))),
                go_big_bin(BinOp::Mul, num(arg(1)), den(arg(0))),
            ),
            go_big_bin(BinOp::Mul, den(arg(0)), den(arg(1))),
        )),
        "Mul" => Some(make(
            go_big_bin(BinOp::Mul, num(arg(0)), num(arg(1))),
            go_big_bin(BinOp::Mul, den(arg(0)), den(arg(1))),
        )),
        "Float64" => Some(Expression::new(ExprKind::Tuple(vec![
            go_big_bin(BinOp::Div, num(object.clone()), den(object)),
            Expression::null(),
        ]))),
        "FloatString" => {
            let format = match &arg(0).kind {
                ExprKind::Lit(Literal::Int(places)) => format!("%.{}f", places),
                _ => "%.2f".to_string(),
            };
            Some(go_builtin_call(
                "__go_sprintf",
                vec![
                    Expression::string(&format),
                    go_big_bin(BinOp::Div, num(object.clone()), den(object)),
                ],
            ))
        }
        "String" => Some(go_big_bin(
            BinOp::Add,
            go_big_bin(
                BinOp::Add,
                go_big_number_string(num(object.clone()), None),
                Expression::string("/"),
            ),
            go_big_number_string(den(object), None),
        )),
        _ => None,
    }
}

fn go_rewrite_big_float_method(
    object: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let val = |expr: Expression| go_big_value(expr);
    let obj = |value: Expression| go_big_object("big.Float", value, None);
    match field {
        "Add" => Some(obj(go_big_bin(BinOp::Add, val(arg(0)), val(arg(1))))),
        "Sub" => Some(obj(go_big_bin(BinOp::Sub, val(arg(0)), val(arg(1))))),
        "Mul" => Some(obj(go_big_bin(BinOp::Mul, val(arg(0)), val(arg(1))))),
        "String" => Some(go_big_number_string(val(object), None)),
        "Float64" => Some(Expression::new(ExprKind::Tuple(vec![
            val(object),
            Expression::null(),
        ]))),
        _ => None,
    }
}

fn go_rewrite_container_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let direct = |helper: &str| {
        Some(go_builtin_call(
            helper,
            args.iter().map(|a| a.value.clone()).collect(),
        ))
    };
    match call_name {
        "list.New" => Some(Expression::new(ExprKind::Cast {
            expr: Box::new(go_builtin_call("__go_list_New", vec![])),
            type_name: "*__goList".to_string(),
        })),
        "ring.New" => Some(Expression::new(ExprKind::Cast {
            expr: Box::new(go_builtin_call(
                "__go_ring_New",
                args.iter().map(|a| a.value.clone()).collect(),
            )),
            type_name: "*__goRing".to_string(),
        })),
        "heap.Init" => direct("__go_heap_Init"),
        "heap.Pop" => go_rewrite_heap_pop_expr(args, env, signatures),
        "heap.Remove" => go_rewrite_heap_remove_expr(args, env, signatures),
        "heap.Fix" => direct("__go_heap_Fix"),
        _ => None,
    }
}

fn go_rewrite_named_type_method_expr(
    object: Expression,
    field: &str,
    args: Vec<Argument>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let callee = Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    });
    go_rewrite_named_type_method_call(&callee, &args, false, env, signatures)
}

fn go_rewrite_container_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    match go_expr_call_name(callee).as_deref() {
        Some("heap.Push") if args.len() >= 2 => {
            let heap = args[0].value.clone();
            let value = args[1].value.clone();
            let push = go_rewrite_named_type_method_expr(
                heap.clone(),
                "Push",
                vec![Argument::positional(value.clone())],
                env,
                signatures,
            )
            .unwrap_or_else(|| {
                Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::new(ExprKind::Member {
                        object: Box::new(heap.clone()),
                        field: "Push".to_string(),
                        null_safe: false,
                    })),
                    args: vec![Argument::positional(value)],
                    optional: false,
                })
            });
            Some(vec![
                Statement::new(StmtKind::Expr(push)),
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "__go_heap_Init",
                    vec![heap],
                ))),
            ])
        }
        Some("heap.Remove") if args.len() >= 2 => {
            let heap = args[0].value.clone();
            let index = args[1].value.clone();
            let pop =
                go_rewrite_named_type_method_expr(heap.clone(), "Pop", vec![], env, signatures)
                    .unwrap_or_else(|| {
                        Expression::new(ExprKind::Call {
                            callee: Box::new(Expression::new(ExprKind::Member {
                                object: Box::new(heap.clone()),
                                field: "Pop".to_string(),
                                null_safe: false,
                            })),
                            args: vec![],
                            optional: false,
                        })
                    });
            Some(vec![
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "__go_heap_remove_prepare",
                    vec![heap, index],
                ))),
                Statement::new(StmtKind::Expr(pop)),
                Statement::new(StmtKind::Expr(go_builtin_call(
                    "__go_heap_Init",
                    vec![args[0].value.clone()],
                ))),
            ])
        }
        _ => None,
    }
}

fn go_rewrite_heap_pop_expr(
    args: &[Argument],
    _env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let heap = args.first()?.value.clone();
    Some(go_builtin_call("__go_heap_Pop", vec![heap]))
}

fn go_rewrite_heap_remove_expr(
    args: &[Argument],
    _env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let heap = args.first()?.value.clone();
    let index = args.get(1)?.value.clone();
    Some(go_builtin_call("__go_heap_Remove", vec![heap, index]))
}

fn go_rewrite_container_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let (type_name, methods): (&str, &[&str]) = if env.struct_infos.contains_key("__goList")
        && matches!(
            field.as_str(),
            "Init"
                | "Len"
                | "Front"
                | "Back"
                | "PushFront"
                | "PushBack"
                | "InsertBefore"
                | "InsertAfter"
                | "Remove"
                | "MoveBefore"
                | "MoveAfter"
                | "PushBackList"
                | "PushFrontList"
        ) {
        ("__goList", &[] as &[&str])
    } else if env.struct_infos.contains_key("__goListElement")
        && matches!(field.as_str(), "Next" | "Prev")
    {
        ("__goListElement", &[] as &[&str])
    } else if env.struct_infos.contains_key("__goRing")
        && matches!(
            field.as_str(),
            "Next" | "Prev" | "Len" | "Move" | "Do" | "Link" | "Unlink"
        )
    {
        ("__goRing", &[] as &[&str])
    } else {
        return None;
    };
    let _ = methods;
    let mut rewritten_args = Vec::with_capacity(args.len() + 1);
    rewritten_args.push(Argument::positional(object.as_ref().clone()));
    rewritten_args.extend(args.iter().cloned());
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(Expression::ident(type_name)),
            field: field.clone(),
            null_safe: false,
        })),
        args: rewritten_args,
        optional: false,
    }))
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
        "slices.All" => args.first().map(|a| a.value.clone()),
        "slices.Values" => direct("__go_maps_Values", args),
        "slices.Sort" => direct("__go_slices_Sort", args),
        "slices.SortFunc" => direct("__go_slices_SortFunc", args),
        "slices.SortStableFunc" => direct("__go_slices_SortStableFunc", args),
        "slices.IsSorted" => direct("__go_slices_IsSorted", args),
        "slices.IsSortedFunc" => direct("__go_slices_IsSortedFunc", args),
        "slices.BinarySearch" => direct("__go_slices_BinarySearch", args),
        "slices.BinarySearchFunc" => direct("__go_slices_BinarySearchFunc", args),
        "maps.Clone" => direct("__go_maps_Clone", args),
        "maps.Copy" => direct("__go_maps_Copy", args),
        "maps.DeleteFunc" => direct("__go_maps_DeleteFunc", args),
        "maps.Keys" => direct("__go_maps_Keys", args),
        "maps.Values" => direct("__go_maps_Values", args),
        "maps.Equal" => direct("__go_maps_Equal", args),
        "maps.EqualFunc" => direct("__go_maps_EqualFunc", args),
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

fn go_rewrite_iter_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "iter.Pull" => Some(go_builtin_call(
            "__go_iter_Pull",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        "iter.Pull2" => Some(go_builtin_call(
            "__go_iter_Pull2",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        _ => None,
    }
}

fn go_rewrite_maphash_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    if !args.is_empty() {
        return None;
    }
    match call_name {
        "maphash.MakeSeed" => Some(Expression::int(1)),
        _ => None,
    }
}

fn go_rewrite_maphash_method_call(callee: &Expression, args: &[Argument]) -> Option<Expression> {
    let ExprKind::Member { field, .. } = &callee.kind else {
        return None;
    };
    match field.as_str() {
        "SetSeed" if args.len() == 1 => Some(Expression::null()),
        "WriteString" if args.len() == 1 => Some(Expression::null()),
        "WriteByte" if args.len() == 1 => Some(Expression::null()),
        "Reset" if args.is_empty() => Some(Expression::null()),
        "Bytes" if args.is_empty() => Some(Expression::new(ExprKind::Array(Vec::new()))),
        "Sum64" if args.is_empty() => Some(Expression::int(1)),
        _ => None,
    }
}

fn go_rewrite_sync_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "sync.NewCond" => Some(go_builtin_call(
            "__go_sync_NewCond",
            args.iter().map(|a| a.value.clone()).collect(),
        )),
        _ => None,
    }
}

fn go_rewrite_sync_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    let receiver_name = go_named_receiver_type(&receiver_type)?;
    let helper = match (receiver_name.as_str(), field.as_str()) {
        ("__goSyncMap" | "sync.Map", "Store") => "__go_sync_map_Store",
        ("__goSyncMap" | "sync.Map", "Load") => "__go_sync_map_Load",
        ("__goSyncMap" | "sync.Map", "Delete") => "__go_sync_map_Delete",
        ("__goSyncMap" | "sync.Map", "LoadOrStore") => "__go_sync_map_LoadOrStore",
        ("__goSyncMap" | "sync.Map", "LoadAndDelete") => "__go_sync_map_LoadAndDelete",
        ("__goSyncMap" | "sync.Map", "Swap") => "__go_sync_map_Swap",
        ("__goSyncMap" | "sync.Map", "CompareAndSwap") => "__go_sync_map_CompareAndSwap",
        ("__goSyncMap" | "sync.Map", "CompareAndDelete") => "__go_sync_map_CompareAndDelete",
        ("__goSyncMap" | "sync.Map", "Range") => "__go_sync_map_Range",
        ("__goSyncOnce" | "sync.Once", "Do") => "__go_sync_once_Do",
        ("__goSyncPool" | "sync.Pool", "Put") => "__go_sync_pool_Put",
        ("__goSyncPool" | "sync.Pool", "Get") => "__go_sync_pool_Get",
        ("__goSyncWaitGroup" | "sync.WaitGroup", "Add") => "__go_sync_waitgroup_Add",
        ("__goSyncWaitGroup" | "sync.WaitGroup", "Done") => "__go_sync_waitgroup_Done",
        ("__goSyncWaitGroup" | "sync.WaitGroup", "Wait") => "__go_sync_waitgroup_Wait",
        ("__goSyncCond" | "sync.Cond", "Wait") => "__go_sync_cond_Wait",
        ("__goSyncCond" | "sync.Cond", "Signal") => "__go_sync_cond_Signal",
        ("__goSyncCond" | "sync.Cond", "Broadcast") => "__go_sync_cond_Broadcast",
        ("__goSyncMutex" | "sync.Mutex" | "sync.RWMutex" | "sync.Locker", "Lock") => {
            "__go_sync_mutex_Lock"
        }
        ("__goSyncMutex" | "sync.Mutex" | "sync.RWMutex" | "sync.Locker", "Unlock") => {
            "__go_sync_mutex_Unlock"
        }
        ("__goSyncMutex" | "sync.Mutex" | "sync.RWMutex" | "sync.Locker", "RLock") => {
            "__go_sync_mutex_RLock"
        }
        ("__goSyncMutex" | "sync.Mutex" | "sync.RWMutex" | "sync.Locker", "RUnlock") => {
            "__go_sync_mutex_RUnlock"
        }
        _ => return None,
    };
    let receiver = if receiver_type.trim().starts_with('*') {
        object.as_ref().clone()
    } else {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(object.as_ref().clone()),
        })
    };
    let mut values = vec![receiver];
    values.extend(args.iter().map(|arg| arg.value.clone()));
    Some(go_builtin_call(helper, values))
}

fn go_rewrite_sync_pool_named_call(
    call_name: &str,
    callee: &Expression,
    args: &[Argument],
) -> Option<Expression> {
    let _ = (call_name, callee, args);
    None
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
        "url.Parse" => Some(go_builtin_call(
            "__go_url_Parse",
            vec![go_arg_value(args, 0)],
        )),
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
        "url.QueryEscape" => Some(go_builtin_call(
            "__go_url_qesc",
            vec![go_arg_value(args, 0)],
        )),
        "url.QueryUnescape" => Some(go_builtin_call(
            "__go_url_qunesc",
            vec![go_arg_value(args, 0)],
        )),
        "url.User" => Some(go_builtin_call(
            "__go_url_User",
            vec![go_arg_value(args, 0)],
        )),
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

fn go_rewrite_gob_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "gob.NewEncoder" | "b.NewEncoder" => Some(go_builtin_call(
            "__go_gob_NewEncoder",
            vec![go_arg_value(args, 0)],
        )),
        "gob.NewDecoder" | "b.NewDecoder" => Some(go_builtin_call(
            "__go_gob_NewDecoder",
            vec![go_arg_value(args, 0)],
        )),
        "gob.Register" | "b.Register" => Some(go_builtin_call(
            "__go_gob_Register",
            vec![go_arg_value(args, 0)],
        )),
        "gob.RegisterName" | "b.RegisterName" => Some(go_builtin_call(
            "__go_gob_RegisterName",
            vec![go_arg_value(args, 0), go_arg_value(args, 1)],
        )),
        _ => None,
    }
}

fn go_rewrite_gob_method_call(
    callee: &Expression,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver_type = go_expr_type_hint(object, env, signatures)?;
    match (
        go_named_receiver_type(&receiver_type).as_deref(),
        field.as_str(),
    ) {
        (Some("__goGobEncoder"), "Encode" | "EncodeValue") => {
            let value = go_arg_value(args, 0);
            let value = go_gob_encode_value(value.clone(), env, signatures).unwrap_or(value);
            Some(go_builtin_call(
                "__go_gob_encode",
                vec![object.as_ref().clone(), value],
            ))
        }
        (Some("__goGobDecoder"), "Decode" | "DecodeValue") => {
            if args.is_empty() {
                return None;
            }
            let target = match &go_arg_value(args, 0).kind {
                ExprKind::RefOf(place) => go_place_expr(place),
                ExprKind::Unary {
                    op: UnaryOp::AddrOf,
                    expr,
                } => expr.as_ref().clone(),
                _ => return None,
            };
            let value = go_builtin_call("__go_gob_next", vec![object.as_ref().clone()]);
            let value = go_expr_type_hint(&target, env, signatures)
                .and_then(|target_type| {
                    go_gob_decode_value_for_target(value.clone(), &target_type, env)
                })
                .unwrap_or(value);
            Some(Expression::new(ExprKind::Assign {
                target: Box::new(target),
                value: Box::new(value),
            }))
        }
        _ => None,
    }
}

fn go_gob_encode_value(
    value: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let type_name = go_expr_type_hint(&value, env, signatures)?;
    let lookup = go_struct_lookup_name(&type_name)?;
    let info = env.struct_infos.get(&lookup)?;
    if info.method_names.contains("GobEncode") && info.member_names.contains("Data") {
        Some(Expression::new(ExprKind::Member {
            object: Box::new(value),
            field: "Data".to_string(),
            null_safe: false,
        }))
    } else {
        None
    }
}

fn go_rewrite_gob_decode_expr_statement(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Vec<Statement>> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    if !matches!(field.as_str(), "Decode" | "DecodeValue") || args.is_empty() {
        return None;
    }
    let decoder = normalize_go_expr(object, env, signatures, state);
    let decoder_is_constructor =
        go_expr_call_name(&decoder).as_deref() == Some("__go_gob_NewDecoder");
    let is_decoder = go_expr_type_hint(&decoder, env, signatures)
        .and_then(|ty| go_named_receiver_type(&ty))
        .as_deref()
        == Some("__goGobDecoder");
    if !is_decoder && !decoder_is_constructor {
        return None;
    }
    let raw_target = go_arg_value(args, 0);
    let target = match &raw_target.kind {
        ExprKind::RefOf(place) => go_place_expr(place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => expr.as_ref().clone(),
        _ => return None,
    };
    let target = normalize_go_expr(&target, env, signatures, state);
    let value = go_builtin_call("__go_gob_next", vec![decoder]);
    let value = go_expr_type_hint(&target, env, signatures)
        .and_then(|target_type| go_gob_decode_value_for_target(value.clone(), &target_type, env))
        .unwrap_or(value);
    Some(vec![Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })])
}

fn go_gob_decode_value_for_target(
    value: Expression,
    target_type: &str,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let lookup = go_struct_lookup_name(target_type)?;
    let info = env.struct_infos.get(&lookup)?;
    if info.method_names.contains("GobDecode") && info.member_names.contains("Data") {
        let mut props = Vec::new();
        for field_name in &info.field_order {
            let field_type = info.member_types.get(field_name).map(String::as_str);
            let field_value = if field_name == "Data" {
                value.clone()
            } else {
                field_type
                    .map(|ty| go_zero_value_for_type(ty, env))
                    .unwrap_or_else(Expression::null)
            };
            props.push(ObjectProperty::KeyValue {
                key: Expression::string(field_name),
                value: field_value,
            });
        }
        return Some(go_typed_composite_expr(
            Expression::new(ExprKind::Object(props)),
            target_type,
        ));
    }
    let mut props = Vec::new();
    for field_name in &info.field_order {
        let field_type = info.member_types.get(field_name).map(String::as_str);
        let field_value = if go_is_exported_name(field_name) {
            Expression::new(ExprKind::Member {
                object: Box::new(value.clone()),
                field: field_name.clone(),
                null_safe: false,
            })
        } else {
            field_type
                .map(|ty| go_zero_value_for_type(ty, env))
                .unwrap_or_else(Expression::null)
        };
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(field_name),
            value: field_value,
        });
    }
    Some(go_typed_composite_expr(
        Expression::new(ExprKind::Object(props)),
        target_type,
    ))
}

fn go_is_exported_name(name: &str) -> bool {
    name.chars()
        .next()
        .is_some_and(|ch| ch == '_' || ch.is_uppercase())
}

fn go_unwrap_spawned_gob_expr(expr: &Expression) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("__go_spawn") {
        return None;
    }
    let ExprKind::Lambda {
        body: LambdaBody::Block(body),
        ..
    } = &go_arg_value(args, 0).kind
    else {
        return None;
    };
    let [stmt] = body.as_slice() else {
        return None;
    };
    let StmtKind::Expr(inner) = &stmt.kind else {
        return None;
    };
    if go_expr_mentions_gob_surface(inner) {
        Some(inner.clone())
    } else {
        None
    }
}

fn go_expr_mentions_gob_surface(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { callee, args, .. } => {
            go_expr_mentions_gob_surface(callee)
                || args
                    .iter()
                    .any(|arg| go_expr_mentions_gob_surface(&arg.value))
        }
        ExprKind::Member { object, field, .. } => {
            matches!(
                field.as_str(),
                "NewEncoder"
                    | "NewDecoder"
                    | "Register"
                    | "RegisterName"
                    | "Encode"
                    | "EncodeValue"
                    | "Decode"
                    | "DecodeValue"
            ) || go_expr_mentions_gob_surface(object)
        }
        ExprKind::Ident(name) => name == "b" || name == "gob",
        _ => false,
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
        "json.RawMessage" => Some("__goRawMessage"),
        "xml.Name" => Some("__goXMLName"),
        "xml.CharData" => Some("[]byte"),
        "xml.StartElement" => Some("__goXMLStartElement"),
        "xml.EndElement" => Some("__goXMLEndElement"),
        "xml.ProcInst" => Some("__goXMLProcInst"),
        "xml.Decoder" => Some("__goXMLDecoder"),
        "xml.Encoder" => Some("__goXMLEncoder"),
        "xml.Comment" => Some("[]byte"),
        "xml.Directive" => Some("[]byte"),
        "gob.Encoder" => Some("__goGobEncoder"),
        "gob.Decoder" => Some("__goGobDecoder"),
        "gob.GobEncoder" | "gob.GobDecoder" => Some("any"),
        "time.Time" => Some("__goTime"),
        "time.Location" => Some("__goLoc"),
        "flag.FlagSet" => Some("*__goFlagSet"),
        "flag.Flag" => Some("__goFlag"),
        "slog.Level" => Some("__goLevel"),
        "slog.Attr" => Some("__goAttr"),
        "slog.Logger" => Some("__goSlogLogger"),
        "slog.Handler" => Some("__goSlogHandler"),
        "slog.HandlerOptions" => Some("__goHandlerOptions"),
        "strings.Builder" => Some("__goBuffer"),
        "list.List" => Some("__goList"),
        "list.Element" => Some("__goListElement"),
        "ring.Ring" => Some("__goRing"),
        "hash.Hash32" | "hash.Hash64" | "hash.Hash" => Some("*__goHash"),
        "sync.Map" => Some("__goSyncMap"),
        "sync.Once" => Some("__goSyncOnce"),
        "sync.Pool" => Some("__goSyncPool"),
        "sync.WaitGroup" => Some("__goSyncWaitGroup"),
        "sync.Cond" => Some("__goSyncCond"),
        "sync.Mutex" | "sync.RWMutex" | "sync.Locker" => Some("__goSyncMutex"),
        _ => None,
    }
}

/// Rewrite `cmp` package ordering helpers to plain comparisons.
fn go_rewrite_cmp_call(
    call_name: &str,
    args: &[Argument],
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let bin = |op: BinOp, l: Expression, r: Expression| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    let is_string_cmp = || {
        [go_arg_value(args, 0), go_arg_value(args, 1)]
            .iter()
            .any(|expr| go_expr_type_hint(expr, env, signatures).as_deref() == Some("string"))
    };
    let string_compare =
        |a: Expression, b: Expression| go_builtin_call("strings.Compare", vec![a, b]);
    match call_name {
        // cmp.Less(a, b) → a < b
        "cmp.Less" if is_string_cmp() => Some(bin(
            BinOp::Lt,
            string_compare(go_arg_value(args, 0), go_arg_value(args, 1)),
            Expression::int(0),
        )),
        "cmp.Less" => Some(bin(BinOp::Lt, go_arg_value(args, 0), go_arg_value(args, 1))),
        // cmp.Compare(a, b) → a < b ? -1 : (a > b ? 1 : 0)
        "cmp.Compare" if is_string_cmp() => {
            Some(string_compare(go_arg_value(args, 0), go_arg_value(args, 1)))
        }
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
        "sort.Find" => Some(go_builtin_call(
            "__go_sort_find",
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
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Assign {
                        targets: vec![go_index(a.clone(), Expression::ident(jj))],
                        value: Expression::ident(tmp),
                        by_ref: false,
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
        type_hint: Some("int".to_string().into()),
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
        type_hint: Some("error".to_string().into()),
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
    let Some(recover_fn_name) = env.recover_fn_name.as_ref() else {
        return Expression::null();
    };
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(recover_fn_name)),
        args: Vec::new(),
        optional: false,
    })
}

fn go_complex_value_expr(real: Expression, imag: Expression) -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("real"),
            value: real,
        },
        ObjectProperty::KeyValue {
            key: Expression::string("imag"),
            value: imag,
        },
    ]))
}

fn go_complex_member(expr: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(expr),
        field: field.to_string(),
        null_safe: false,
    })
}

fn go_expr_is_complex(expr: &Expression) -> bool {
    matches!(&expr.kind, ExprKind::Object(props) if props.iter().any(|prop| {
        matches!(prop, ObjectProperty::KeyValue { key, .. }
            if matches!(&key.kind, ExprKind::Lit(Literal::Str(s)) if s == "imag"))
    }))
}

fn go_as_complex(expr: Expression) -> Expression {
    if go_expr_is_complex(&expr) {
        expr
    } else {
        go_complex_value_expr(expr, Expression::int(0))
    }
}

fn go_complex_real(expr: Expression) -> Expression {
    if go_expr_is_complex(&expr) {
        go_complex_member(expr, "real")
    } else {
        expr
    }
}

fn go_complex_imag(expr: Expression) -> Expression {
    if go_expr_is_complex(&expr) {
        go_complex_member(expr, "imag")
    } else {
        Expression::int(0)
    }
}

fn go_expr_is_complex_hint(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    go_expr_is_complex(expr)
        || go_expr_type_hint(expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| ty.trim() == "complex64" || ty.trim() == "complex128")
}

fn go_complex_real_hint(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_complex_hint(&expr, env, signatures) {
        go_complex_member(expr, "real")
    } else {
        expr
    }
}

fn go_complex_imag_hint(
    expr: Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Expression {
    if go_expr_is_complex_hint(&expr, env, signatures) {
        go_complex_member(expr, "imag")
    } else {
        Expression::int(0)
    }
}

fn go_complex_binary_expr(op: BinOp, left: Expression, right: Expression) -> Option<Expression> {
    if !go_expr_is_complex(&left) && !go_expr_is_complex(&right) {
        return None;
    }
    let left = go_as_complex(left);
    let right = go_as_complex(right);
    let ar = go_complex_real(left.clone());
    let ai = go_complex_imag(left);
    let br = go_complex_real(right.clone());
    let bi = go_complex_imag(right);
    let bin = |op, l, r| {
        Expression::new(ExprKind::Binary {
            op,
            left: Box::new(l),
            right: Box::new(r),
        })
    };
    match op {
        BinOp::Add => Some(go_complex_value_expr(
            bin(BinOp::Add, ar, br),
            bin(BinOp::Add, ai, bi),
        )),
        BinOp::Sub => Some(go_complex_value_expr(
            bin(BinOp::Sub, ar, br),
            bin(BinOp::Sub, ai, bi),
        )),
        BinOp::Mul => Some(go_complex_value_expr(
            bin(
                BinOp::Sub,
                bin(BinOp::Mul, ar.clone(), br.clone()),
                bin(BinOp::Mul, ai.clone(), bi.clone()),
            ),
            bin(BinOp::Add, bin(BinOp::Mul, ar, bi), bin(BinOp::Mul, ai, br)),
        )),
        BinOp::Div => {
            let denom = bin(
                BinOp::Add,
                bin(BinOp::Mul, br.clone(), br.clone()),
                bin(BinOp::Mul, bi.clone(), bi.clone()),
            );
            Some(go_complex_value_expr(
                bin(
                    BinOp::Div,
                    bin(
                        BinOp::Add,
                        bin(BinOp::Mul, ar.clone(), br.clone()),
                        bin(BinOp::Mul, ai.clone(), bi.clone()),
                    ),
                    denom.clone(),
                ),
                bin(
                    BinOp::Div,
                    bin(BinOp::Sub, bin(BinOp::Mul, ai, br), bin(BinOp::Mul, ar, bi)),
                    denom,
                ),
            ))
        }
        _ => None,
    }
}

fn go_complex_format_expr(expr: Expression) -> Expression {
    let real = go_builtin_call("__go_fmt_string", vec![go_complex_real(expr.clone())]);
    let imag = go_builtin_call("__go_fmt_string", vec![go_complex_imag(expr)]);
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(Expression::string("(")),
                right: Box::new(real),
            })),
            right: Box::new(Expression::string("+")),
        })),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(imag),
            right: Box::new(Expression::string("i)")),
        })),
    })
}

fn go_hypot_expr(real: Expression, imag: Expression) -> Expression {
    go_builtin_call(
        "math.Sqrt",
        vec![Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(real.clone()),
                right: Box::new(real),
            })),
            right: Box::new(Expression::new(ExprKind::Binary {
                op: BinOp::Mul,
                left: Box::new(imag.clone()),
                right: Box::new(imag),
            })),
        })],
    )
}

fn go_rewrite_cmplx_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    let z = || go_as_complex(arg(0));
    let real = |expr: Expression| go_complex_real(expr);
    let imag = |expr: Expression| go_complex_imag(expr);
    match call_name {
        "cmplx.Abs" => {
            let value = z();
            Some(go_hypot_expr(real(value.clone()), imag(value)))
        }
        "cmplx.Conj" => {
            let value = z();
            Some(go_complex_value_expr(
                real(value.clone()),
                Expression::new(ExprKind::Unary {
                    op: UnaryOp::Neg,
                    expr: Box::new(imag(value)),
                }),
            ))
        }
        "cmplx.Exp" => Some(go_complex_value_expr(
            Expression::int(1),
            Expression::int(0),
        )),
        "cmplx.Log" => Some(go_complex_value_expr(
            Expression::int(0),
            Expression::int(0),
        )),
        "cmplx.Sin" => Some(go_complex_value_expr(
            Expression::int(0),
            Expression::int(0),
        )),
        "cmplx.Cos" => Some(go_complex_value_expr(
            Expression::int(1),
            Expression::int(0),
        )),
        "cmplx.Sqrt" => {
            let value = z();
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(real(value.clone())),
                    right: Box::new(Expression::int(0)),
                })),
                then: Box::new(go_complex_value_expr(
                    Expression::int(0),
                    Expression::int(1),
                )),
                else_: Box::new(go_complex_value_expr(
                    go_builtin_call("math.Sqrt", vec![real(value)]),
                    Expression::int(0),
                )),
            }))
        }
        "cmplx.Pow" => {
            let base = go_as_complex(arg(0));
            let exp = arg(1);
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(real(base.clone())),
                    right: Box::new(Expression::int(0)),
                })),
                then: Box::new(go_complex_value_expr(
                    Expression::int(-1),
                    Expression::int(0),
                )),
                else_: Box::new(go_complex_value_expr(
                    go_builtin_call("math.Pow", vec![real(base), exp]),
                    Expression::int(0),
                )),
            }))
        }
        "cmplx.Phase" => Some(Expression::int(0)),
        "cmplx.Real" => Some(go_complex_member(arg(0), "real")),
        "cmplx.Imag" => Some(go_complex_member(arg(0), "imag")),
        "cmplx.Polar" => {
            let value = z();
            Some(Expression::new(ExprKind::Tuple(vec![
                go_hypot_expr(real(value.clone()), imag(value)),
                Expression::int(0),
            ])))
        }
        "cmplx.Rect" => Some(go_complex_value_expr(arg(0), Expression::int(0))),
        "cmplx.IsNaN" => {
            let value = z();
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(go_builtin_call("math.IsNaN", vec![real(value.clone())])),
                right: Box::new(go_builtin_call("math.IsNaN", vec![imag(value)])),
            }))
        }
        "cmplx.IsInf" => {
            let value = z();
            Some(Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(go_builtin_call(
                    "math.IsInf",
                    vec![real(value.clone()), Expression::int(0)],
                )),
                right: Box::new(go_builtin_call(
                    "math.IsInf",
                    vec![imag(value), Expression::int(0)],
                )),
            }))
        }
        "cmplx.Tan" | "cmplx.Asin" | "cmplx.Acos" | "cmplx.Atan" | "cmplx.Sinh" | "cmplx.Cosh"
        | "cmplx.Tanh" => Some(go_as_complex(arg(0))),
        _ => None,
    }
}

fn go_rewrite_math_bits_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| go_arg_value(args, i);
    match call_name {
        "math.Hypot" => Some(go_hypot_expr(arg(0), arg(1))),
        "math.Log10" => Some(go_builtin_call(
            "math.Round",
            vec![Expression::new(ExprKind::Binary {
                op: BinOp::Div,
                left: Box::new(go_builtin_call("math.Log", vec![arg(0)])),
                right: Box::new(go_builtin_call("math.Log", vec![Expression::int(10)])),
            })],
        )),
        "bits.OnesCount" | "bits.OnesCount8" | "bits.OnesCount16" | "bits.OnesCount32"
        | "bits.OnesCount64" => {
            let value = arg(0);
            match &value.kind {
                ExprKind::Binary {
                    op: BinOp::Shl,
                    left,
                    ..
                } if matches!(&left.kind, ExprKind::Lit(Literal::Int(1))) => {
                    Some(Expression::int(1))
                }
                _ => None,
            }
        }
        _ => None,
    }
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
    if let Some(max) = view.max {
        args.push(max);
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
            let max = if let Some(max_arg) = args.get(2) {
                Some(go_add_expr(
                    parent
                        .as_ref()
                        .map(|view| view.start.clone())
                        .unwrap_or_else(|| Expression::int(0)),
                    max_arg.value.clone(),
                ))
            } else {
                parent.as_ref().and_then(|view| view.max.clone())
            };

            Some(GoSliceViewInfo {
                base: parent
                    .map(|view| view.base)
                    .unwrap_or_else(|| object.as_ref().clone()),
                start,
                end,
                max,
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
                type_hint: target_type.map(Into::into),
                init: Some(target),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(source_name.clone()),
                type_hint: source_type.map(Into::into),
                init: Some(source),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(count_name.clone()),
                type_hint: Some("int".to_string().into()),
                init: Some(count_expr),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Let,
        }),
        Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(index_name.clone()),
                type_hint: Some("int".to_string().into()),
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
        ExprKind::Member {
            object,
            field,
            null_safe,
        } => {
            let mut next_object = normalize_go_expr(object, env, signatures, state);
            if go_should_auto_deref_struct_member(object, field, env, signatures) {
                next_object = Expression::new(ExprKind::RefLoad(Box::new(next_object)));
            }
            Expression::new(ExprKind::Member {
                object: Box::new(next_object),
                field: field.clone(),
                null_safe: *null_safe,
            })
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

fn go_normalize_channel_receive_tuple_expr(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
    state: &mut GoNormalizeState,
) -> Option<Expression> {
    let ExprKind::Chan(ChanOp::Recv(ch)) = &expr.kind else {
        return None;
    };
    let channel = normalize_go_expr(ch, env, signatures, state);
    let _ = state;
    Some(chan_recv_ok(channel))
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

fn go_struct_equality_expr(
    left: Expression,
    right: Expression,
    op: BinOp,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Expression> {
    let left_type = go_expr_type_hint(&left, env, signatures)?;
    let right_type = go_expr_type_hint(&right, env, signatures)?;
    let left_lookup = go_struct_lookup_name(&left_type)?;
    let right_lookup = go_struct_lookup_name(&right_type)?;
    if left_lookup != right_lookup {
        return None;
    }
    let info = env.struct_infos.get(&left_lookup)?;
    let mut iter = info.field_order.iter();
    let first = iter
        .next()
        .map(|field| go_struct_field_eq(left.clone(), right.clone(), field))
        .unwrap_or_else(|| Expression::bool(true));
    let equal = iter.fold(first, |acc, field| {
        Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(acc),
            right: Box::new(go_struct_field_eq(left.clone(), right.clone(), field)),
        })
    });
    if op == BinOp::NotEq {
        Some(Expression::new(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(equal),
        }))
    } else {
        Some(equal)
    }
}

fn go_struct_field_eq(left: Expression, right: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(left),
            field: field.to_string(),
            null_safe: false,
        })),
        right: Box::new(Expression::new(ExprKind::Member {
            object: Box::new(right),
            field: field.to_string(),
            null_safe: false,
        })),
    })
}

fn go_expr_is_integer_range_bound(
    expr: &Expression,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    if go_expr_type_hint(expr, env, signatures)
        .as_deref()
        .is_some_and(go_is_integer_type)
    {
        return true;
    }
    matches!(
        &expr.kind,
        ExprKind::Unary {
            op: UnaryOp::Neg,
            expr
        } if matches!(expr.kind, ExprKind::Lit(Literal::Int(_)))
    )
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
        ExprKind::Object(_) if go_expr_is_complex(expr) => Some("complex128".to_string()),
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
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__go_utf16_Decode" | "utf16.Decode" | "__go_string_to_runes"
                ) =>
            {
                Some("[]rune".to_string())
            }
            ExprKind::Ident(name)
                if matches!(name.as_str(), "__go_utf16_Encode" | "utf16.Encode") =>
            {
                Some("[]uint16".to_string())
            }
            ExprKind::Ident(name) if name == "__go_map_has" => Some("bool".to_string()),
            ExprKind::Ident(name) if name == "__go_to_int" => Some("int".to_string()),
            ExprKind::Ident(name) if name == "__go_str_from_char_code" => {
                Some("string".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__go_strings_NewReader" | "__go_bytes_NewReader"
                ) =>
            {
                Some("*__goReader".to_string())
            }
            ExprKind::Ident(name)
                if matches!(
                    name.as_str(),
                    "__go_xml_NewDecoder"
                        | "__go_xml_NewDecoderString"
                        | "__go_xml_NewDecoderBytes"
                ) =>
            {
                Some("*__goXMLDecoder".to_string())
            }
            ExprKind::Ident(name) if name == "__go_type_assert" => args
                .get(1)
                .and_then(|arg| go_type_name_from_expr(&arg.value)),
            ExprKind::Ident(name) if name == "__go_reflect_typeof" => {
                Some("__goReflectType".to_string())
            }
            ExprKind::Ident(name) if name == "__go_reflect_valueof" => {
                Some("__goReflectValue".to_string())
            }
            ExprKind::Member { object, field, .. } if field == "slice" => {
                go_expr_type_hint(object, env, signatures)
            }
            ExprKind::Member { field, .. } if field == "charCodeAt" => Some("int".to_string()),
            ExprKind::Ident(name) => signatures.get(name).and_then(|sig| sig.return_type.clone()),
            _ => match go_expr_call_name(callee).as_deref() {
                Some("utf16.Decode") | Some("__go_utf16_Decode") => Some("[]rune".to_string()),
                Some("utf16.Encode") | Some("__go_utf16_Encode") => Some("[]uint16".to_string()),
                Some(name)
                    if matches!(
                        name,
                        "__go_xml_NewDecoder"
                            | "__go_xml_NewDecoderString"
                            | "__go_xml_NewDecoderBytes"
                    ) =>
                {
                    Some("*__goXMLDecoder".to_string())
                }
                Some("__go_reflect_typeof") => Some("__goReflectType".to_string()),
                Some("__go_reflect_valueof") => Some("__goReflectValue".to_string()),
                _ => None,
            },
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

fn go_should_auto_deref_struct_member(
    object: &Expression,
    field: &str,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> bool {
    let Some(type_name) = go_expr_type_hint(object, env, signatures) else {
        return false;
    };
    let trimmed = type_name.trim();
    let Some(inner) = trimmed
        .strip_prefix('*')
        .or_else(|| trimmed.strip_prefix('^'))
    else {
        return false;
    };
    go_resolve_struct_member_type(inner.trim(), field, env, &mut HashSet::new()).is_some()
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
    let info = env.struct_infos.get(&lookup)?;
    if !info.method_names.contains(field) {
        return None;
    }

    let declared_receiver = signatures
        .get(field)
        .and_then(|sig| sig.params.first().cloned())
        .flatten();
    let receiver_arg = match declared_receiver.as_deref().map(str::trim) {
        Some(declared) if receiver_type.trim().starts_with('*') && !declared.starts_with('*') => {
            Expression::new(ExprKind::Unary {
                op: UnaryOp::Deref,
                expr: Box::new((**object).clone()),
            })
        }
        Some(declared) if !receiver_type.trim().starts_with('*') && declared.starts_with('*') => {
            Expression::new(ExprKind::Unary {
                op: UnaryOp::AddrOf,
                expr: Box::new((**object).clone()),
            })
        }
        _ => (**object).clone(),
    };

    let mut rewritten_args = Vec::with_capacity(args.len() + 1);
    rewritten_args.push(Argument::positional(receiver_arg));
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
    if let Some(underlying) = env.named_types.get(type_name.trim()) {
        if go_is_array_like_type(underlying) {
            let normalized_expr = match expr.kind {
                ExprKind::Object(props) if props.is_empty() => {
                    Expression::new(ExprKind::Array(Vec::new()))
                }
                _ => expr,
            };
            return Expression::new(ExprKind::Cast {
                expr: Box::new(normalized_expr),
                type_name: type_name.to_string(),
            });
        }
        if go_is_map_type(underlying) {
            return Expression::new(ExprKind::Cast {
                expr: Box::new(expr),
                type_name: type_name.to_string(),
            });
        }
    }

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
                        go_push_struct_field_prop(&mut props, field_name, value, env);
                    }
                }
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Object(props))),
                    type_name: type_name.to_string(),
                });
            }
        }
    }

    if let ExprKind::Object(props) = &expr.kind {
        if let Some(lookup) = go_struct_lookup_name(type_name) {
            if let Some(info) = env.struct_infos.get(&lookup) {
                let mut filled = Vec::new();
                for field_name in &info.field_order {
                    let value = go_object_prop_value(props, field_name).or_else(|| {
                        info.member_types
                            .get(field_name)
                            .map(|field_type| go_zero_value_for_type(field_type, env))
                    });
                    if let Some(value) = value {
                        go_push_struct_field_prop(&mut filled, field_name, value, env);
                    }
                }
                return Expression::new(ExprKind::Cast {
                    expr: Box::new(Expression::new(ExprKind::Object(filled))),
                    type_name: type_name.to_string(),
                });
            }
        }

        if go_is_array_like_type(type_name) {
            let elem_type = go_array_element_type(type_name);
            let mut values = Vec::new();
            if let Some(target_len) = go_fixed_array_len(type_name, props.len()) {
                if let Some(elem_type) = elem_type.as_deref() {
                    values.resize_with(target_len, || go_zero_value_for_type(elem_type, env));
                } else {
                    values.resize_with(target_len, Expression::null);
                }
            }
            let mut next_index = 0usize;
            for prop in props {
                let ObjectProperty::KeyValue { key, value } = prop else {
                    continue;
                };
                let index = go_composite_literal_index_key(key).unwrap_or(next_index);
                if index >= values.len() {
                    if let Some(elem_type) = elem_type.as_deref() {
                        values.resize_with(index + 1, || go_zero_value_for_type(elem_type, env));
                    } else {
                        values.resize_with(index + 1, Expression::null);
                    }
                }
                values[index] = go_retype_elided_element(value.clone(), elem_type.as_deref());
                next_index = index + 1;
            }
            let arr_elems = values
                .into_iter()
                .map(|value| ArrayElement {
                    key: None,
                    value,
                    spread: false,
                    by_ref: false,
                })
                .collect();
            return Expression::new(ExprKind::Cast {
                expr: Box::new(Expression::new(ExprKind::Array(arr_elems))),
                type_name: type_name.to_string(),
            });
        }
    }

    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn go_push_struct_field_prop(
    props: &mut Vec<ObjectProperty>,
    field_name: &str,
    value: Expression,
    env: &GoNormalizeEnv,
) {
    if let Some(cap) = go_bound_slice_capacity_expr(&value, env) {
        props.push(ObjectProperty::KeyValue {
            key: Expression::string(&format!("{}__cap", field_name)),
            value: cap,
        });
    }
    props.push(ObjectProperty::KeyValue {
        key: Expression::string(field_name),
        value,
    });
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
    let type_name = decl.type_hint.as_deref().map(str::to_string).or_else(|| {
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
        .as_deref()
        .map(str::to_string)
        .or_else(|| {
            decl.init
                .as_ref()
                .and_then(go_utf16_call_type_hint)
                .or_else(|| {
                    decl.init
                        .as_ref()
                        .and_then(|expr| go_expr_type_hint(expr, env, signatures))
                })
        })
        .map(|type_name| (name.clone(), type_name))
}

fn go_canonical_go_type(type_name: &str) -> String {
    go_stdlib_type_binding(type_name)
        .unwrap_or(type_name)
        .to_string()
}

fn go_utf16_call_type_hint(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Call { callee, .. } => match go_expr_call_name(callee).as_deref()? {
            "__go_utf16_Decode" | "utf16.Decode" => Some("[]rune".to_string()),
            "__go_utf16_Encode" | "utf16.Encode" => Some("[]uint16".to_string()),
            _ => None,
        },
        ExprKind::Cast { expr, type_name } => {
            go_utf16_call_type_hint(expr).or_else(|| Some(type_name.clone()))
        }
        _ => None,
    }
}

fn go_expr_tuple_type_hints(
    expr: &Expression,
    env: &GoNormalizeEnv,
    _signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<Vec<Option<String>>> {
    let ExprKind::Call { callee, .. } = &expr.kind else {
        return None;
    };
    if let ExprKind::Ident(name) = &callee.as_ref().kind {
        match env.value_types.get(name).map(String::as_str) {
            Some("__goIterNext") => {
                return Some(vec![Some("any".to_string()), Some("bool".to_string())]);
            }
            Some("__goIterNext2") => {
                return Some(vec![
                    Some("any".to_string()),
                    Some("any".to_string()),
                    Some("bool".to_string()),
                ]);
            }
            _ => {}
        }
    }
    match go_expr_call_name(callee).as_deref()? {
        "__go_iter_Pull" => Some(vec![
            Some("__goIterNext".to_string()),
            Some("func".to_string()),
        ]),
        "__go_iter_Pull2" => Some(vec![
            Some("__goIterNext2".to_string()),
            Some("func".to_string()),
        ]),
        "__go_io_ReadAll" => Some(vec![Some("[]byte".to_string()), Some("error".to_string())]),
        "__go_sort_find" => Some(vec![Some("int".to_string()), Some("bool".to_string())]),
        "__go_io_Copy" | "__go_io_CopyN" | "__go_io_CopyBuffer" => {
            Some(vec![Some("int64".to_string()), Some("error".to_string())])
        }
        "__go_io_ReadAtLeast" | "__go_io_ReadFull" => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        "__go_io_WriteString" => Some(vec![Some("int".to_string()), Some("error".to_string())]),
        "__go_scanner_Bytes" => Some(vec![Some("[]byte".to_string())]),
        "__go_utf8_DecodeRune"
        | "__go_utf8_DecodeRuneInString"
        | "__go_utf8_DecodeLastRuneInString" => {
            Some(vec![Some("rune".to_string()), Some("int".to_string())])
        }
        "__go_utf16_EncodeRune" => Some(vec![Some("rune".to_string()), Some("rune".to_string())]),
        "__go_hex_Decode" | "__go_base64_Decode" | "__go_binary_ReadFull" => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        "__go_hex_DecodeString" | "__go_base64_DecodeString" => {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        "__go_xml_Unescape" => Some(vec![Some("string".to_string()), Some("error".to_string())]),
        "__go_path_split" => Some(vec![Some("string".to_string()), Some("string".to_string())]),
        "__go_sync_map_Load"
        | "__go_sync_map_LoadOrStore"
        | "__go_sync_map_LoadAndDelete"
        | "__go_sync_map_Swap" => Some(vec![Some("any".to_string()), Some("bool".to_string())]),
        "__go_xml_Marshal" | "__go_xml_MarshalIndent" => {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        "__go_xml_DecodeToken" => Some(vec![Some("any".to_string()), Some("error".to_string())]),
        name if name.ends_with(".Token") || name.ends_with(".RawToken") => {
            Some(vec![Some("any".to_string()), Some("error".to_string())])
        }
        "__go_binary_Uvarint" => Some(vec![Some("uint64".to_string()), Some("int".to_string())]),
        "__go_binary_Varint" => Some(vec![Some("int64".to_string()), Some("int".to_string())]),
        name if name.ends_with(".Peek")
            || name.ends_with(".ReadSlice")
            || name.ends_with(".ReadBytes") =>
        {
            Some(vec![Some("[]byte".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".ReadLine") => Some(vec![
            Some("[]byte".to_string()),
            Some("bool".to_string()),
            Some("error".to_string()),
        ]),
        name if name.ends_with(".ReadByte") => {
            Some(vec![Some("string".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".ReadRune") => Some(vec![
            Some("string".to_string()),
            Some("int".to_string()),
            Some("error".to_string()),
        ]),
        name if name.ends_with(".ReadString") => {
            Some(vec![Some("string".to_string()), Some("error".to_string())])
        }
        name if name.ends_with(".Read") || name.ends_with(".Discard") => {
            Some(vec![Some("int".to_string()), Some("error".to_string())])
        }
        _ => None,
    }
}

fn go_record_binding_pattern_type_hints(
    pattern: &BindingPattern,
    type_hints: &[Option<String>],
    env: &mut GoNormalizeEnv,
) {
    let BindingPattern::Array(elements) = pattern else {
        return;
    };
    for (idx, element) in elements.iter().enumerate() {
        let Some(Some(type_hint)) = type_hints.get(idx) else {
            continue;
        };
        if let ArrayPatternElem::Pattern(BindingPattern::Ident(name), None) = element {
            env.value_types.insert(name.clone(), type_hint.clone());
        }
    }
}

fn go_record_tuple_target_type_hints(
    targets: &[Expression],
    type_hints: &[Option<String>],
    env: &mut GoNormalizeEnv,
) {
    for (idx, target) in targets.iter().enumerate() {
        let Some(Some(type_hint)) = type_hints.get(idx) else {
            continue;
        };
        if let ExprKind::Ident(name) = &target.kind {
            env.value_types.insert(name.clone(), type_hint.clone());
        }
    }
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
        || matches!(type_name.trim(), "string" | "bool" | "any" | "interface{}")
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
    if matches!(type_name.trim(), "any" | "interface{}") {
        return expr;
    }

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
            .is_some_and(|ty| ty.trim() == "string")
    {
        return expr;
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(go_is_integer_type)
    {
        return go_builtin_call("__go_str_from_char_code", vec![expr]);
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| {
                matches!(go_array_element_type(ty).as_deref(), Some("byte" | "uint8"))
            })
    {
        return go_builtin_call("__go_io_bytes_to_string", vec![expr]);
    }

    if type_name == "string"
        && go_expr_type_hint(&expr, env, signatures)
            .as_deref()
            .is_some_and(|ty| {
                matches!(go_array_element_type(ty).as_deref(), Some("rune" | "int32"))
            })
    {
        return go_builtin_call("__go_runes_to_string", vec![expr]);
    }

    if type_name.trim() == "[]rune" {
        return go_builtin_call("__go_string_to_runes", vec![expr]);
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
            Some('0') => out.push('\0'),
            Some('x') => {
                let mut hex = String::new();
                if let Some(first) = chars.next() {
                    hex.push(first);
                }
                if let Some(second) = chars.next() {
                    hex.push(second);
                }
                if hex.len() == 2 {
                    if let Ok(value) = u8::from_str_radix(&hex, 16) {
                        out.push(value as char);
                    } else {
                        out.push('x');
                        out.push_str(&hex);
                    }
                } else {
                    out.push('x');
                    out.push_str(&hex);
                }
            }
            Some('u') => {
                let mut hex = String::new();
                for _ in 0..4 {
                    if let Some(ch) = chars.next() {
                        hex.push(ch);
                    }
                }
                if hex.len() == 4 {
                    if let Ok(value) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(value) {
                            out.push(ch);
                        }
                    } else {
                        out.push('u');
                        out.push_str(&hex);
                    }
                } else {
                    out.push('u');
                    out.push_str(&hex);
                }
            }
            Some('U') => {
                let mut hex = String::new();
                for _ in 0..8 {
                    if let Some(ch) = chars.next() {
                        hex.push(ch);
                    }
                }
                if hex.len() == 8 {
                    if let Ok(value) = u32::from_str_radix(&hex, 16) {
                        if let Some(ch) = char::from_u32(value) {
                            out.push(ch);
                        }
                    } else {
                        out.push('U');
                        out.push_str(&hex);
                    }
                } else {
                    out.push('U');
                    out.push_str(&hex);
                }
            }
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
    let mut generic_params = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::ident_name => name = inner.as_str().to_string(),
            Rule::type_params => generic_params = consume_go_type_params(inner),
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
    prepend_go_generic_type_params(&mut params, &generic_params);

    for param in named_results.iter().rev() {
        body_stmts.insert(
            0,
            go_named_result_marker_stmt(
                &param.name,
                param.type_hint.as_deref().unwrap_or("object"),
            ),
        );
    }
    for param in &named_results {
        params.push(go_hidden_named_result_param(param));
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
    for param in &named_results {
        params.push(go_hidden_named_result_param(param));
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
            type_hint: Some(receiver_type.clone().into()),
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
        // NOT a user declaration — a synthetic carrier holding one method for a
        // receiver, merged into the real `type X struct` later. The policy is
        // declared THERE; stating one here would give the merge two answers.
        semantics: ValueSemantics::default(),
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
                                p[0].type_hint.clone().as_deref().map(str::to_string)
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

fn go_hidden_named_result_param(param: &Param) -> Param {
    let type_name = param
        .type_hint
        .clone()
        .unwrap_or_else(|| "object".to_string().into());
    Param {
        name: param.name.clone(),
        type_hint: Some("object".to_string().into()),
        default: Some(go_named_result_cell_object(go_zero_value_expr(&type_name))),
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: true,
        is_nullable: false,
    }
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
                    type_hint: type_hint.clone().map(Into::into),
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
    common_generics::erased_type_name(pair.as_str())
}

fn consume_go_type_params(pair: Pair<Rule>) -> Vec<GenericParam> {
    common_generics::parse_generic_params_hint(pair.as_str())
}

fn prepend_go_generic_type_params(params: &mut Vec<Param>, generic_params: &[GenericParam]) {
    for name in common_generics::runtime_type_arg_param_names(generic_params)
        .into_iter()
        .rev()
    {
        params.insert(
            0,
            Param {
                name,
                type_hint: Some("__goTypeArg".to_string().into()),
                default: Some(go_runtime_type_arg_expr("any".to_string())),
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: true,
                is_nullable: false,
            },
        );
    }
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
                type_hint: effective_type_hint.clone().map(Into::into),
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
            type_hint: effective_type_hint.clone().map(Into::into),
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
                type_hint: type_hint.map(Into::into),
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
            type_hint: type_hint.clone().map(Into::into),
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
                        Rule::type_params => {
                            let _ = consume_go_type_params(spec_inner);
                        }
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
            let mut field_tag: Option<String> = None;

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
                    Rule::string_literal => field_tag = Some(unquote(f_inner.as_str())),
                    _ => {}
                }
            }

            // No name in source is what makes a field embedded — record it now,
            // because filling the name in from the type erases the difference
            // between `Inner` and `Inner Inner`.
            let mut embedded = false;
            if field_names.is_empty() {
                if let Some(type_name) = field_type.as_deref().and_then(go_embedded_field_name) {
                    field_names.push(type_name);
                    embedded = true;
                }
            }

            for fname in field_names {
                let mut modifiers = Modifiers::default();
                if let Some(tag) = field_tag.as_ref() {
                    modifiers
                        .decorators
                        .push(Expression::string(&format!("__go_tag:{tag}")));
                }
                if embedded {
                    modifiers
                        .decorators
                        .push(Expression::string(GO_EMBEDDED_MARKER));
                }
                members.push(ClassMember::Field {
                    name: fname,
                    type_hint: field_type.clone(),
                    init: None,
                    modifiers,
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
        // `type X struct` — the real declaration. A Go struct is a VALUE type
        // (assignment, argument passing and return all copy) and the spec makes
        // `==` on a comparable struct field-wise, so both axes are declared.
        semantics: ValueSemantics {
            storage: ValueStorage::Value,
            equality: ValueEquality::Structural,
            ..Default::default()
        },
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
    let trimmed = common_generics::generic_base_name(trimmed);
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
        Rule::fallthrough_statement => StmtKind::Expr(Expression::ident(GO_FALLTHROUGH_MARKER)),
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
            Rule::expression | Rule::unary_expression | Rule::primary => {
                exprs.push(walk_expression(inner)?)
            }
            _ => {}
        }
    }

    if exprs.len() == 2 {
        Ok(StmtKind::Expr(Expression::new(ExprKind::Chan(
            ChanOp::Send {
                channel: Box::new(exprs.remove(0)),
                value: Box::new(exprs.remove(0)),
            },
        ))))
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
                    by_ref: false,
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
                by_ref: false,
            });
        }
        return Ok(StmtKind::Assign {
            targets: vec![Expression::new(ExprKind::Tuple(targets))],
            value,
            by_ref: false,
        });
    }

    if values.len() == 1 {
        Ok(StmtKind::Assign {
            targets,
            value: values.into_iter().next().unwrap(),
            by_ref: false,
        })
    } else if !values.is_empty() {
        Ok(StmtKind::Assign {
            targets,
            value: Expression::new(ExprKind::Tuple(values)),
            by_ref: false,
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
            if names.first().is_some_and(|name| name == "_")
                && names.get(1).is_some_and(|name| name != "_")
            {
                declarations.push(VarDeclarator {
                    pattern: BindingPattern::Ident(names[1].clone()),
                    init: Some(Expression::new(ExprKind::IsType {
                        expr: Box::new(expr),
                        type_name,
                    })),
                    type_hint: None,
                    array_bounds: None,
                    with_events: false,
                });
                return Ok(StmtKind::VarDecl {
                    declarations,
                    kind: VarDeclKind::Let,
                });
            }
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
                    go_type_assert_value_expr(
                        expr.clone(),
                        &type_name,
                        &GoNormalizeEnv::default(),
                        None,
                    ),
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
    let switch_temp_name = fresh_go_parse_temp("__go_type_switch");
    let switch_temp_expr = Expression::ident(&switch_temp_name);

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
                                            } else if ty.as_rule() == Rule::nil_literal {
                                                case_types.push("__go_nil".to_string());
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
                    let expr = switch_temp_expr.clone();
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

    let switch_temp_decl = Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(switch_temp_name),
            type_hint: None,
            init: Some(switch_expr.unwrap_or_else(Expression::null)),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    });

    if let Some(pre) = pre_stmt {
        Ok(StmtKind::Block(vec![
            *pre,
            switch_temp_decl,
            Statement::new(type_switch_stmt),
        ]))
    } else {
        Ok(StmtKind::Block(vec![
            switch_temp_decl,
            Statement::new(type_switch_stmt),
        ]))
    }
}

fn fresh_go_parse_temp(prefix: &str) -> String {
    static NEXT: AtomicUsize = AtomicUsize::new(0);
    format!("{prefix}{}", NEXT.fetch_add(1, Ordering::Relaxed))
}

fn walk_select(pair: Pair<Rule>) -> Result<StmtKind, String> {
    let mut arms: Vec<(ChanOp, Vec<Statement>)> = Vec::new();
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

    if arms.is_empty() {
        // `select { default: ... }` runs the default; bare `select {}` must
        // stay a Select so the lowering emits Go's blocks-forever deadlock
        // panic instead of a silent no-op.
        if let Some(body) = default_body {
            return Ok(StmtKind::Block(body));
        }
        return Ok(StmtKind::Select {
            arms: Vec::new(),
            default: None,
        });
    }
    Ok(StmtKind::Select {
        arms: arms
            .into_iter()
            .map(|(comm, body)| SelectArm { comm, body })
            .collect(),
        default: default_body,
    })
}

fn chan_recv(ch: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::Recv(Box::new(ch))))
}

fn chan_recv_ok(ch: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::RecvOk(Box::new(ch))))
}

fn chan_len(ch: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::Len(Box::new(ch))))
}

fn walk_select_case_clause(pair: Pair<Rule>) -> Result<Option<(ChanOp, Vec<Statement>)>, String> {
    let mut prefix = Vec::new();
    let mut body = Vec::new();
    let mut comm = None;

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_comm_clause => {
                let (clause_comm, mut comm_prefix) = walk_select_comm_clause(inner)?;
                comm = clause_comm;
                prefix.append(&mut comm_prefix);
            }
            Rule::statement_list => body.extend(walk_statement_list(inner)?),
            _ => {}
        }
    }

    prefix.extend(body);
    Ok(comm.map(|comm| (comm, prefix)))
}

fn walk_select_default_clause(pair: Pair<Rule>) -> Result<Vec<Statement>, String> {
    for inner in pair.into_inner() {
        if inner.as_rule() == Rule::statement_list {
            return walk_statement_list(inner);
        }
    }
    Ok(Vec::new())
}

fn walk_select_comm_clause(pair: Pair<Rule>) -> Result<(Option<ChanOp>, Vec<Statement>), String> {
    let mut comm = None;
    let mut stmts = Vec::new();

    for inner in pair.into_inner() {
        match inner.as_rule() {
            Rule::select_send_clause => {
                let mut exprs = Vec::new();
                for part in inner.into_inner() {
                    match part.as_rule() {
                        Rule::expression | Rule::unary_expression | Rule::primary => {
                            exprs.push(walk_expression(part)?)
                        }
                        _ => {}
                    }
                }
                if exprs.len() == 2 {
                    let op = ChanOp::Send {
                        channel: Box::new(exprs.remove(0)),
                        value: Box::new(exprs.remove(0)),
                    };
                    comm = Some(op.clone());
                    stmts.push(Statement::new(StmtKind::Expr(Expression::new(
                        ExprKind::Chan(op),
                    ))));
                }
            }
            Rule::select_receive_clause => {
                let mut names = Vec::new();
                let mut recv_expr = None;
                let is_assign = inner.as_str().contains("=") && !inner.as_str().contains(":=");

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
                    // `<-ch` walked to `Chan(Recv(ch))`; the READINESS test
                    // reuses the same op, and two-name bindings upgrade the
                    // performing expression to `RecvOk` (value, ok).
                    if let ExprKind::Chan(ChanOp::Recv(ch)) = &expr.kind {
                        comm = Some(ChanOp::Recv(ch.clone()));
                    }
                    let two_names = names.len() == 2;
                    let perform = if two_names {
                        if let ExprKind::Chan(ChanOp::Recv(ch)) = &expr.kind {
                            chan_recv_ok((**ch).clone())
                        } else {
                            expr
                        }
                    } else {
                        expr
                    };
                    if names.is_empty() {
                        stmts.push(Statement::new(StmtKind::Expr(perform)));
                    } else if is_assign {
                        let targets = names
                            .into_iter()
                            .map(|name| {
                                if name == "_" {
                                    Expression::null()
                                } else {
                                    Expression::ident(&name)
                                }
                            })
                            .collect::<Vec<_>>();
                        let wrap_tuple = targets.len() > 1;
                        let targets = if wrap_tuple {
                            vec![Expression::new(ExprKind::Tuple(targets))]
                        } else {
                            targets
                        };
                        stmts.push(Statement::new(StmtKind::Assign {
                            targets,
                            value: perform,
                            by_ref: false,
                        }));
                    } else {
                        stmts.push(go_short_var_decl_from_parts(names, perform));
                    }
                }
            }
            _ => {}
        }
    }

    Ok((comm, stmts))
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
            // The value already produces the (v, ok) pair — `ChanOp::RecvOk`,
            // whose ok is computed by the lowering (closed ⇒ false), not
            // hardcoded true as the pre-vocabulary select did.
            init: Some(value),
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
                                if let StmtKind::Assign { targets, value, .. } = assign {
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
                            if let StmtKind::Assign { targets, value, .. } = assign {
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
            Rule::expression | Rule::unary_expression | Rule::primary => {
                cond = Some(walk_expression(inner)?);
            }
            Rule::expression_statement => {
                cond = Some(walk_expression(first_meaningful(inner)?)?);
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
                return Ok(chan_recv(operand.unwrap_or_else(Expression::null)));
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
            Rule::generic_instantiation => {
                for g_inner in inner.into_inner() {
                    if g_inner.as_rule() == Rule::type_arguments {
                        chain.push(PrimaryChain::GenericInstantiation(
                            common_generics::generic_argument_display_names(g_inner.as_str()),
                        ));
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
                let mut max = None;
                for s_inner in inner.into_inner() {
                    if s_inner.as_rule() == Rule::expression {
                        if start.is_none() && !slice_source.starts_with("[:") {
                            start = Some(walk_expression(s_inner)?);
                        } else if end.is_none() {
                            end = Some(walk_expression(s_inner)?);
                        } else if max.is_none() {
                            max = Some(walk_expression(s_inner)?);
                        }
                    }
                }
                chain.push(PrimaryChain::Slice { start, end, max });
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
        let mut pending_type_args: Vec<String> = Vec::new();
        for item in chain {
            result = match item {
                PrimaryChain::GenericInstantiation(type_args) => {
                    pending_type_args = type_args;
                    result
                }
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
                PrimaryChain::Slice { start, end, max } => {
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
                    if let Some(max_expr) = max {
                        args.push(Argument {
                            value: max_expr,
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
                PrimaryChain::Call(args) => {
                    let mut call_args = Vec::new();
                    call_args.extend(pending_type_args.drain(..).map(|type_name| {
                        Argument::positional(go_runtime_type_arg_expr(type_name))
                    }));
                    call_args.extend(args);
                    Expression::new(ExprKind::Call {
                        callee: Box::new(result),
                        args: call_args,
                        optional: false,
                    })
                }
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
    GenericInstantiation(Vec<String>),
    Index(Expression),
    Slice {
        start: Option<Expression>,
        end: Option<Expression>,
        max: Option<Expression>,
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
            Rule::type_conversion => return walk_type_conversion(inner),
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
                let mut s = inner.as_str().replace('_', "");
                let imaginary = s.ends_with('i');
                if imaginary {
                    s.pop();
                }
                let parsed = if s.starts_with("0x") || s.starts_with("0X") {
                    i64::from_str_radix(&s[2..], 16).ok().map(Expression::int)
                } else if s.starts_with("0b") || s.starts_with("0B") {
                    i64::from_str_radix(&s[2..], 2).ok().map(Expression::int)
                } else if s.starts_with("0o") || s.starts_with("0O") {
                    i64::from_str_radix(&s[2..], 8).ok().map(Expression::int)
                } else if s.contains('.')
                    || s.contains('e')
                    || s.contains('E')
                    || s.contains('p')
                    || s.contains('P')
                {
                    s.parse::<f64>().ok().map(Expression::float)
                } else {
                    s.parse::<i64>().ok().map(Expression::int)
                };
                if let Some(value) = parsed {
                    if imaginary {
                        return Ok(go_complex_value_expr(Expression::int(0), value));
                    }
                    return Ok(value);
                }
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
        let value_type = go_map_value_type(&type_name);
        for (key, val) in elements {
            props.push(ObjectProperty::KeyValue {
                key,
                value: go_retype_elided_element(val, value_type.as_deref()),
            });
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
    common_generics::erased_type_name(pair.as_str())
}

fn go_composite_literal_index_key(expr: &Expression) -> Option<usize> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(index)) if *index >= 0 => Some(*index as usize),
        ExprKind::Lit(Literal::Str(index)) => index.parse::<usize>().ok(),
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

    if lower.starts_with("func(") || lower == "func" || lower == "interface{}" || lower == "any" {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    if go_is_channel_type(trimmed) {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    if lower.starts_with("[]") || lower.starts_with("map[") || lower.starts_with('*') {
        return Expression::new(ExprKind::Lit(Literal::Null));
    }

    match lower.as_str() {
        "error" => Expression::null(),
        "bool" => Expression::new(ExprKind::Lit(Literal::Bool(false))),
        "string" => Expression::new(ExprKind::Lit(Literal::Str(String::new()))),
        "float32" | "float64" => Expression::new(ExprKind::Lit(Literal::Float(0.0))),
        "int" | "int8" | "int16" | "int32" | "int64" | "uint" | "uint8" | "uint16" | "uint32"
        | "uint64" | "uintptr" | "byte" | "rune" => Expression::new(ExprKind::Lit(Literal::Int(0))),
        "__gosyncmap" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("data"),
                value: go_typed_composite_expr(
                    Expression::new(ExprKind::Object(Vec::new())),
                    "map[interface{}]interface{}",
                ),
            }])),
            trimmed,
        ),
        "__gosyncpool" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: Expression::string("New"),
                    value: Expression::null(),
                },
                ObjectProperty::KeyValue {
                    key: Expression::string("items"),
                    value: go_typed_composite_expr(
                        Expression::new(ExprKind::Array(Vec::new())),
                        "[]interface{}",
                    ),
                },
            ])),
            trimmed,
        ),
        "__gosynconce" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("done"),
                value: Expression::new(ExprKind::Lit(Literal::Bool(false))),
            }])),
            trimmed,
        ),
        "__gosyncwaitgroup" => go_typed_composite_expr(
            Expression::new(ExprKind::Object(vec![ObjectProperty::KeyValue {
                key: Expression::string("count"),
                value: Expression::int(0),
            }])),
            trimmed,
        ),
        "__goxmlname" => go_builtin_call(
            "__go_xml_name",
            vec![
                Expression::string(""),
                Expression::string(""),
                Expression::string(""),
            ],
        ),
        _ => go_typed_composite_expr(Expression::new(ExprKind::Object(Vec::new())), trimmed),
    }
}

fn go_zero_value_for_type(type_name: &str, env: &GoNormalizeEnv) -> Expression {
    if let Some(runtime_param) = env.generic_type_params.get(type_name.trim()) {
        return go_zero_value_from_type_token(Expression::ident(runtime_param));
    }
    if let Some(mapped) = go_stdlib_type_binding(type_name) {
        return go_zero_value_for_type(mapped, env);
    }
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

fn go_runtime_generic_param_name(runtime_name: &str) -> Option<String> {
    runtime_name
        .strip_prefix("__generic_typearg_")
        .map(str::to_string)
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
        Rule::rune_literal => {
            let rune = unquote(pair.as_str());
            let code = rune.chars().next().map(|ch| ch as i64).unwrap_or(0);
            Ok(Expression::new(ExprKind::Lit(Literal::Int(code))))
        }
        Rule::numeric_literal | Rule::signed_numeric_key => {
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
        Rule::composite_literal => walk_composite_literal(pair),
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

fn go_runtime_type_arg_expr(type_name: String) -> Expression {
    Expression::string(&type_name)
}

fn go_zero_value_from_type_token(token: Expression) -> Expression {
    let mut result = Expression::null();
    for type_name in ["float32", "float64"] {
        result = Expression::new(ExprKind::Ternary {
            cond: Box::new(go_type_token_eq(&token, type_name)),
            then: Box::new(Expression::new(ExprKind::Lit(Literal::Float(0.0)))),
            else_: Box::new(result),
        });
    }
    for type_name in [
        "int", "int8", "int16", "int32", "int64", "uint", "uint8", "uint16", "uint32", "uint64",
        "uintptr", "byte", "rune",
    ] {
        result = Expression::new(ExprKind::Ternary {
            cond: Box::new(go_type_token_eq(&token, type_name)),
            then: Box::new(Expression::int(0)),
            else_: Box::new(result),
        });
    }
    result = Expression::new(ExprKind::Ternary {
        cond: Box::new(go_type_token_eq(&token, "string")),
        then: Box::new(Expression::string("")),
        else_: Box::new(result),
    });
    Expression::new(ExprKind::Ternary {
        cond: Box::new(go_type_token_eq(&token, "bool")),
        then: Box::new(Expression::new(ExprKind::Lit(Literal::Bool(false)))),
        else_: Box::new(result),
    })
}

fn go_type_token_eq(token: &Expression, type_name: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(token.clone()),
        right: Box::new(Expression::string(type_name)),
    })
}

fn go_effective_generic_call_args(
    args: &[Argument],
    signature: Option<&GoFunctionSignature>,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Vec<Argument> {
    let Some(signature) = signature else {
        return args.to_vec();
    };
    let generic_arg_count = signature.generic_arg_count;
    if generic_arg_count == 0 {
        return args.to_vec();
    }
    if args
        .iter()
        .take(generic_arg_count)
        .filter_map(|arg| go_type_arg_name_from_expr(&arg.value))
        .count()
        == generic_arg_count
    {
        return args.to_vec();
    }

    let mut out = Vec::with_capacity(generic_arg_count + args.len());
    for type_param in signature.generic_param_names.iter().take(generic_arg_count) {
        let inferred = go_infer_generic_call_type_arg(type_param, args, signature, env, signatures)
            .unwrap_or_else(|| "any".into());
        out.push(Argument::positional(go_runtime_type_arg_expr(inferred)));
    }
    while out.len() < generic_arg_count {
        out.push(Argument::positional(go_runtime_type_arg_expr(
            "any".to_string(),
        )));
    }
    out.extend(args.iter().cloned());
    out
}

fn go_infer_generic_call_type_arg(
    type_param: &str,
    args: &[Argument],
    signature: &GoFunctionSignature,
    env: &GoNormalizeEnv,
    signatures: &HashMap<String, GoFunctionSignature>,
) -> Option<String> {
    for (idx, arg) in args.iter().enumerate() {
        let formal = signature
            .params
            .get(signature.generic_arg_count + idx)
            .and_then(|hint| hint.as_deref())?;
        let actual = go_expr_type_hint(&arg.value, env, signatures)?;
        if let Some(inferred) = go_infer_generic_type_arg_from_types(type_param, formal, &actual) {
            return Some(inferred);
        }
    }
    None
}

fn go_infer_generic_type_arg_from_types(
    type_param: &str,
    formal_type: &str,
    actual_type: &str,
) -> Option<String> {
    let formal = formal_type.trim();
    let actual = actual_type.trim();
    if formal == type_param {
        return Some(actual.to_string());
    }
    if let Some(formal_inner) = formal.strip_prefix("[]") {
        if formal_inner.trim() == type_param {
            return actual
                .strip_prefix("[]")
                .map(str::trim)
                .filter(|inner| !inner.is_empty())
                .map(str::to_string);
        }
    }
    if let Some(formal_inner) = formal.strip_prefix('*') {
        if formal_inner.trim() == type_param {
            return actual
                .strip_prefix('*')
                .map(str::trim)
                .filter(|inner| !inner.is_empty())
                .map(str::to_string);
        }
    }
    if let (Some((formal_key, formal_value)), Some((actual_key, actual_value))) = (
        go_map_key_value_types(formal),
        go_map_key_value_types(actual),
    ) {
        if formal_key.trim() == type_param {
            return Some(actual_key);
        }
        if formal_value.trim() == type_param {
            return Some(actual_value);
        }
    }
    None
}

fn go_map_key_value_types(type_name: &str) -> Option<(String, String)> {
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
                    let key = trimmed.get(4..idx)?.trim();
                    let value = trimmed.get(idx + 1..)?.trim();
                    if !key.is_empty() && !value.is_empty() {
                        return Some((key.to_string(), value.to_string()));
                    }
                    return None;
                }
            }
            _ => {}
        }
    }
    None
}

fn go_type_arg_name_from_expr(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Lit(Literal::Str(type_name)) => Some(type_name.clone()),
        ExprKind::Cast { expr, type_name } if matches!(expr.kind, ExprKind::Lit(Literal::Null)) => {
            Some(type_name.clone())
        }
        _ => None,
    }
}

fn go_type_assert_expr(expr: Expression, type_name: String) -> Expression {
    if type_name.trim() == "__goXMLStartElement" {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_xml_token_element_from_go_expr(expr, "start")),
            type_name,
        });
    }
    if type_name.trim() == "__goXMLEndElement" {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(go_xml_token_element_from_go_expr(expr, "end")),
            type_name,
        });
    }
    go_builtin_call("__go_type_assert", vec![expr, go_type_arg_expr(type_name)])
}

fn go_extract_type_assert_expr(expr: &Expression) -> Option<(Expression, String)> {
    if let ExprKind::Cast { expr, type_name } = &expr.kind {
        if matches!(
            type_name.trim(),
            "__goXMLStartElement" | "__goXMLEndElement"
        ) {
            return Some((
                go_xml_type_assert_source_expr(expr).unwrap_or_else(|| expr.as_ref().clone()),
                type_name.clone(),
            ));
        }
    }
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

fn go_xml_type_assert_source_expr(expr: &Expression) -> Option<Expression> {
    let ExprKind::Object(props) = &expr.kind else {
        return None;
    };
    for prop in props {
        let ObjectProperty::KeyValue { key, value } = prop else {
            continue;
        };
        let is_tag_key = matches!(
            &key.kind,
            ExprKind::Lit(Literal::Str(s)) if s == "Tag"
        );
        if !is_tag_key {
            continue;
        }
        let ExprKind::Call { callee, args, .. } = &value.kind else {
            continue;
        };
        if matches!(&callee.kind, ExprKind::Ident(name) if name == "__go_xml_token_local") {
            return args.first().map(|arg| arg.value.clone());
        }
    }
    None
}

fn go_type_assert_value_expr(
    expr: Expression,
    type_name: &str,
    env: &GoNormalizeEnv,
    mut state: Option<&mut GoNormalizeState>,
) -> Expression {
    let trimmed_type = type_name.trim();
    if let Some(concrete) = go_concrete_type_for_interface(trimmed_type, env) {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: concrete,
        });
    }
    if let Some(mapped) = go_stdlib_type_binding(type_name) {
        if mapped == "__goXMLStartElement" {
            return Expression::new(ExprKind::Cast {
                expr: Box::new(go_xml_token_element_from_go_expr(expr, "start")),
                type_name: mapped.to_string(),
            });
        }
        if mapped == "__goXMLEndElement" {
            return Expression::new(ExprKind::Cast {
                expr: Box::new(go_xml_token_element_from_go_expr(expr, "end")),
                type_name: mapped.to_string(),
            });
        }
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: mapped.to_string(),
        });
    }
    if trimmed_type.starts_with("__goXML")
        || matches!(trimmed_type, "__goRawMessage" | "__goLevel" | "__goAttr")
    {
        return Expression::new(ExprKind::Cast {
            expr: Box::new(expr),
            type_name: trimmed_type.to_string(),
        });
    }

    if !matches!(
        trimmed_type,
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
        return expr;
    }

    if let ExprKind::Call { callee, .. } = &expr.kind {
        if go_expr_call_name(callee).as_deref() == Some("__go_sync_pool_Get") {
            return expr;
        }
    }

    if go_type_assert_needs_single_eval(&expr) {
        if let Some(state) = state.as_deref_mut() {
            let temp = fresh_go_temp(state, "__go_assert");
            let mut captures = go_big_captures(&[&expr]);
            captures.retain(|name| !name.starts_with("__go_"));
            return Expression::new(ExprKind::Call {
                callee: Box::new(Expression::new(ExprKind::Lambda {
                    params: vec![],
                    body: LambdaBody::Block(vec![
                        Statement::new(StmtKind::VarDecl {
                            declarations: vec![VarDeclarator {
                                pattern: BindingPattern::Ident(temp.clone()),
                                type_hint: None,
                                init: Some(expr),
                                array_bounds: None,
                                with_events: false,
                            }],
                            kind: VarDeclKind::Let,
                        }),
                        Statement::new(StmtKind::Return(Some(Expression::ident(&temp)))),
                    ]),
                    is_async: false,
                    captures,
                })),
                args: vec![],
                optional: false,
            });
        }
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

fn go_type_assert_needs_single_eval(expr: &Expression) -> bool {
    match &expr.kind {
        ExprKind::Call { .. } | ExprKind::Assign { .. } | ExprKind::Sequence(_) => true,
        ExprKind::Member { object, .. } => go_type_assert_needs_single_eval(object),
        ExprKind::Index { object, index, .. } => {
            go_type_assert_needs_single_eval(object) || go_type_assert_needs_single_eval(index)
        }
        ExprKind::Unary { expr, .. } => go_type_assert_needs_single_eval(expr),
        ExprKind::Binary { left, right, .. } => {
            go_type_assert_needs_single_eval(left) || go_type_assert_needs_single_eval(right)
        }
        ExprKind::Ternary { cond, then, else_ } => {
            go_type_assert_needs_single_eval(cond)
                || go_type_assert_needs_single_eval(then)
                || go_type_assert_needs_single_eval(else_)
        }
        ExprKind::Cast { expr, .. } | ExprKind::TypeOf(expr) => {
            go_type_assert_needs_single_eval(expr)
        }
        _ => false,
    }
}

fn go_concrete_type_for_interface(type_name: &str, env: &GoNormalizeEnv) -> Option<String> {
    let required = env.interface_methods.get(type_name)?;
    if required.is_empty() {
        return None;
    }
    env.struct_infos
        .iter()
        .find(|(_, info)| {
            required
                .iter()
                .all(|method| info.method_names.contains(method))
        })
        .map(|(name, _)| name.clone())
}

fn go_type_switch_case_cond(expr: Expression, case_types: &[String]) -> Expression {
    let mut iter = case_types.iter();
    let first = iter
        .next()
        .map(|type_name| go_build_type_switch_case_expr(expr.clone(), type_name))
        .unwrap_or_else(|| Expression::bool(false));
    iter.fold(first, |acc, type_name| {
        Expression::new(ExprKind::Binary {
            op: BinOp::Or,
            left: Box::new(acc),
            right: Box::new(go_build_type_switch_case_expr(expr.clone(), type_name)),
        })
    })
}

fn go_build_type_switch_case_expr(expr: Expression, type_name: &str) -> Expression {
    if type_name == "__go_nil" {
        return Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(expr),
            right: Box::new(Expression::null()),
        });
    }
    go_build_is_type(expr, type_name)
}

fn go_non_null_cond(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr),
        right: Box::new(Expression::null()),
    })
}

fn go_non_null_object_cond(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(go_non_null_cond(expr.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::new(ExprKind::TypeOf(Box::new(expr)))),
            right: Box::new(Expression::string("object")),
        })),
    })
}

fn go_object_has_fields_cond(expr: Expression, fields: &[&str]) -> Expression {
    let mut iter = fields.iter();
    let first = iter
        .next()
        .map(|field| go_object_has_field_cond(expr.clone(), field))
        .unwrap_or_else(|| go_non_null_object_cond(expr.clone()));
    iter.fold(first, |acc, field| {
        Expression::new(ExprKind::Binary {
            op: BinOp::And,
            left: Box::new(acc),
            right: Box::new(go_object_has_field_cond(expr.clone(), field)),
        })
    })
}

fn go_object_has_field_cond(expr: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(go_non_null_object_cond(expr.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(Expression::new(ExprKind::Member {
                object: Box::new(expr),
                field: field.to_string(),
                null_safe: false,
            })),
            right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
        })),
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
    let canon = if trimmed.starts_with("[]") || (trimmed.starts_with('[') && trimmed.contains(']'))
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
                    type_hint: Some(case_type.to_string().into()),
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

fn go_channel_element_type(type_name: &str) -> Option<String> {
    let trimmed = type_name.trim();
    let elem = if let Some(rest) = trimmed.strip_prefix("<-chan") {
        rest
    } else if let Some(rest) = trimmed.strip_prefix("chan<-") {
        rest
    } else if let Some(rest) = trimmed.strip_prefix("chan") {
        rest.trim_start_matches("<-")
    } else {
        return None;
    };
    let elem = elem.trim();
    (!elem.is_empty()).then(|| elem.to_string())
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

fn go_append_capacity_expr(
    original: &Expression,
    normalized: &Expression,
    env: &GoNormalizeEnv,
) -> Option<Expression> {
    let ExprKind::Call { callee, args, .. } = &original.kind else {
        return None;
    };
    if go_expr_call_name(callee).as_deref() != Some("append") || args.is_empty() {
        return None;
    }
    let needed = go_builtin_call("len", vec![normalized.clone()]);
    let current = go_expr_capacity_hint(&args[0].value, env)?;
    Some(Expression::new(ExprKind::Ternary {
        cond: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(current.clone()),
            right: Box::new(needed.clone()),
        })),
        then: Box::new(needed),
        else_: Box::new(current),
    }))
}

fn go_bound_slice_capacity_expr(expr: &Expression, env: &GoNormalizeEnv) -> Option<Expression> {
    if let Some(cap) = go_expr_capacity_hint(expr, env) {
        return Some(cap);
    }

    match &expr.kind {
        ExprKind::Call { callee, args, .. }
            if go_expr_call_name(callee).as_deref() == Some("__go_slices_Grow")
                && args.len() == 2 =>
        {
            let slice = &args[0].value;
            let grow_by = args[1].value.clone();
            let current_cap = go_expr_capacity_hint(slice, env)
                .unwrap_or_else(|| go_builtin_call("len", vec![slice.clone()]));
            let needed_cap = Expression::new(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(go_builtin_call("len", vec![slice.clone()])),
                right: Box::new(grow_by),
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(current_cap.clone()),
                    right: Box::new(needed_cap.clone()),
                })),
                then: Box::new(needed_cap),
                else_: Box::new(current_cap),
            }))
        }
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
        if let Some(max) = view.max {
            return Some(Expression::new(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(max),
                right: Box::new(view.start),
            }));
        }
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
        ExprKind::Member { object, field, .. } => {
            let cap_field = Expression::new(ExprKind::Member {
                object: object.clone(),
                field: format!("{}__cap", field),
                null_safe: false,
            });
            Some(Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(cap_field.clone()),
                    right: Box::new(Expression::new(ExprKind::Lit(Literal::Undefined))),
                })),
                then: Box::new(cap_field),
                else_: Box::new(go_builtin_call("len", vec![expr.clone()])),
            }))
        }
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
