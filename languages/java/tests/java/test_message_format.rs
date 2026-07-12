use crate::helpers::run_main;

#[test]
fn message_format_single_positional_argument() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("Hello, {0}!"); System.out.println(mf.format(new Object[]{"World"}));"#,
    );
    assert_eq!(out, vec!["Hello, World!"]);
}

#[test]
fn message_format_two_positional_arguments_in_order() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0} bought {1} apples"); System.out.println(mf.format(new Object[]{"Alice", 6}));"#,
    );
    assert_eq!(out, vec!["Alice bought 6 apples"]);
}

#[test]
fn message_format_reversed_placeholder_indices() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{1} then {0}"); System.out.println(mf.format(new Object[]{"second", "first"}));"#,
    );
    assert_eq!(out, vec!["first then second"]);
}

#[test]
fn message_format_reuses_same_argument_index() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0} plus {0} equals {1}"); System.out.println(mf.format(new Object[]{3, 6}));"#,
    );
    assert_eq!(out, vec!["3 plus 3 equals 6"]);
}

#[test]
fn message_format_three_arguments_mixed_types() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("user={0}, score={1}, active={2}"); System.out.println(mf.format(new Object[]{"bob", 99, true}));"#,
    );
    assert_eq!(out, vec!["user=bob, score=99, active=true"]);
}

#[test]
fn message_format_static_format_method() {
    let out = run_main(
        r#"System.out.println(java.text.MessageFormat.format("Total: {0}", new Object[]{42}));"#,
    );
    assert_eq!(out, vec!["Total: 42"]);
}

#[test]
fn message_format_choice_zero_branch() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#no files|1#one file|1<many files}"); System.out.println(mf.format(new Object[]{0}));"#,
    );
    assert_eq!(out, vec!["no files"]);
}

#[test]
fn message_format_choice_one_branch() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#no files|1#one file|1<many files}"); System.out.println(mf.format(new Object[]{1}));"#,
    );
    assert_eq!(out, vec!["one file"]);
}

#[test]
fn message_format_choice_many_branch() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#no files|1#one file|1<many files}"); System.out.println(mf.format(new Object[]{5}));"#,
    );
    assert_eq!(out, vec!["many files"]);
}

#[test]
fn message_format_choice_upper_bound_exact() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#empty|1#single|2#pair|2<lots}"); System.out.println(mf.format(new Object[]{2}));"#,
    );
    assert_eq!(out, vec!["pair"]);
}

#[test]
fn message_format_choice_negative_index() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,-1#below zero|0#zero|0<positive}"); System.out.println(mf.format(new Object[]{-3}));"#,
    );
    assert_eq!(out, vec!["below zero"]);
}

#[test]
fn message_format_number_subformat_integer() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("count={0,number,integer}", java.util.Locale.US); System.out.println(mf.format(new Object[]{12345.67}));"#,
    );
    assert_eq!(out, vec!["count=12,345"]);
}

#[test]
fn message_format_number_subformat_currency() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("price={0,number,currency}", java.util.Locale.US); System.out.println(mf.format(new Object[]{19.99}));"#,
    );
    assert_eq!(out, vec!["price=$19.99"]);
}

#[test]
fn message_format_number_subformat_percent() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("rate={0,number,percent}", java.util.Locale.US); System.out.println(mf.format(new Object[]{0.75}));"#,
    );
    assert_eq!(out, vec!["rate=75%"]);
}

#[test]
fn message_format_number_subformat_custom_pattern() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("val={0,number,0.00}", java.util.Locale.US); System.out.println(mf.format(new Object[]{3.1}));"#,
    );
    assert_eq!(out, vec!["val=3.10"]);
}

#[test]
fn message_format_date_subformat_short_style() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 5, 15, 0, 0, 0); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("due {0,date,short}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["due 6/15/24"]);
}

#[test]
fn message_format_date_subformat_custom_pattern() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 5, 15, 0, 0, 0); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("on {0,date,yyyy-MM-dd}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["on 2024-06-15"]);
}

#[test]
fn message_format_time_subformat_short_style() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 0, 1, 14, 30, 0); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("at {0,time,short}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["at 2:30 PM"]);
}

#[test]
fn message_format_time_subformat_medium_style() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 0, 1, 9, 5, 7); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("clock {0,time,medium}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["clock 9:05:07 AM"]);
}

#[test]
fn message_format_time_subformat_custom_hms() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 0, 1, 8, 9, 10); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("t={0,time,HH:mm:ss}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["t=08:09:10"]);
}

#[test]
fn message_format_escape_single_quote_doubles_it() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("it''s {0}"); System.out.println(mf.format(new Object[]{"fine"}));"#,
    );
    assert_eq!(out, vec!["it's fine"]);
}

#[test]
fn message_format_literal_text_between_placeholders() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("from {0} to {1} inclusive"); System.out.println(mf.format(new Object[]{1, 10}));"#,
    );
    assert_eq!(out, vec!["from 1 to 10 inclusive"]);
}

#[test]
fn message_format_empty_string_argument() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("name=''{0}''"); System.out.println(mf.format(new Object[]{""}));"#,
    );
    assert_eq!(out, vec!["name=''"]);
}

#[test]
fn message_format_null_argument_prints_null() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("value={0}"); System.out.println(mf.format(new Object[]{null}));"#,
    );
    assert_eq!(out, vec!["value=null"]);
}

#[test]
fn message_format_choice_and_number_combined() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#no items|1#one item|1<{0,number,integer} items}", java.util.Locale.US); System.out.println(mf.format(new Object[]{12}));"#,
    );
    assert_eq!(out, vec!["12 items"]);
}

#[test]
fn message_format_choice_and_string_combined() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#Hello {1}|1#Hi {1}|1<Hey {1}}"); System.out.println(mf.format(new Object[]{0, "team"})); System.out.println(mf.format(new Object[]{2, "team"}));"#,
    );
    assert_eq!(out, vec!["Hello team", "Hey team"]);
}

#[test]
fn message_format_apply_pattern_changes_output() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("old {0}"); mf.applyPattern("new {0}"); System.out.println(mf.format(new Object[]{"x"}));"#,
    );
    assert_eq!(out, vec!["new x"]);
}

#[test]
fn message_format_to_pattern_returns_applied_pattern() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("a {0} b"); System.out.println(mf.toPattern());"#,
    );
    assert_eq!(out, vec!["a {0} b"]);
}

#[test]
fn message_format_set_locale_affects_number_formatting() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("n={0,number,integer}"); mf.setLocale(java.util.Locale.US); System.out.println(mf.format(new Object[]{1000}));"#,
    );
    assert_eq!(out, vec!["n=1,000"]);
}

#[test]
fn message_format_parse_simple_positional_message() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("score={0}"); Object[] parsed = mf.parse("score=42"); System.out.println(parsed[0]);"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn message_format_parse_two_argument_message() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0}:{1}"); Object[] parsed = mf.parse("alpha:beta"); System.out.println(parsed[0]); System.out.println(parsed[1]);"#,
    );
    assert_eq!(out, vec!["alpha", "beta"]);
}

#[test]
fn message_format_format_parse_roundtrip_simple() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("id={0}, name={1}"); Object[] args = new Object[]{7, "vybe"}; String s = mf.format(args); Object[] parsed = mf.parse(s); System.out.println(parsed[0]); System.out.println(parsed[1]);"#,
    );
    assert_eq!(out, vec!["7", "vybe"]);
}

#[test]
fn message_format_integer_via_number_style() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("n={0,number}", java.util.Locale.US); System.out.println(mf.format(new Object[]{500}));"#,
    );
    assert_eq!(out, vec!["n=500"]);
}

#[test]
fn message_format_currency_in_sentence() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0} paid {1,number,currency}", java.util.Locale.US); System.out.println(mf.format(new Object[]{"Sam", 12.5}));"#,
    );
    assert_eq!(out, vec!["Sam paid $12.50"]);
}

#[test]
fn message_format_date_and_time_same_argument() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 5, 15, 14, 30, 0); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("on {0,date,yyyy-MM-dd} at {0,time,HH:mm}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["on 2024-06-15 at 14:30"]);
}

#[test]
fn message_format_multiple_number_formats_in_one_pattern() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("a={0,number,integer} b={1,number,0.0}", java.util.Locale.US); System.out.println(mf.format(new Object[]{1000, 2.25}));"#,
    );
    assert_eq!(out, vec!["a=1,000 b=2.3"]);
}

#[test]
fn message_format_choice_exact_boundary_match() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#zero|1#one|1<more}"); System.out.println(mf.format(new Object[]{1}));"#,
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn message_format_choice_default_greater_than_branch() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#zero|1#one|1<more}"); System.out.println(mf.format(new Object[]{99}));"#,
    );
    assert_eq!(out, vec!["more"]);
}

#[test]
fn message_format_argument_index_at_end_of_string() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("answer: {0}"); System.out.println(mf.format(new Object[]{"yes"}));"#,
    );
    assert_eq!(out, vec!["answer: yes"]);
}

#[test]
fn message_format_zero_items_plural_style_choice() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("You have {0,choice,0#no messages|1#one message|1<{0,number,integer} messages}", java.util.Locale.US); System.out.println(mf.format(new Object[]{0}));"#,
    );
    assert_eq!(out, vec!["You have no messages"]);
}

#[test]
fn message_format_one_item_plural_style_choice() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("You have {0,choice,0#no messages|1#one message|1<{0,number,integer} messages}", java.util.Locale.US); System.out.println(mf.format(new Object[]{1}));"#,
    );
    assert_eq!(out, vec!["You have one message"]);
}

#[test]
fn message_format_many_items_plural_style_with_embedded_number() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("You have {0,choice,0#no messages|1#one message|1<{0,number,integer} messages}", java.util.Locale.US); System.out.println(mf.format(new Object[]{42}));"#,
    );
    assert_eq!(out, vec!["You have 42 messages"]);
}

#[test]
fn message_format_escaped_curly_braces_as_literal() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("set {0} = '{'value'}'"); System.out.println(mf.format(new Object[]{"x"}));"#,
    );
    assert_eq!(out, vec!["set x = {value}"]);
}

#[test]
fn message_format_date_subformat_full_style() {
    let out = run_main(
        r#"java.util.GregorianCalendar cal = new java.util.GregorianCalendar(java.util.TimeZone.getTimeZone("UTC"), java.util.Locale.US); cal.set(2024, 0, 1, 0, 0, 0); cal.set(java.util.Calendar.MILLISECOND, 0); java.text.MessageFormat mf = new java.text.MessageFormat("{0,date,full}", java.util.Locale.US); System.out.println(mf.format(new Object[]{cal.getTime()}));"#,
    );
    assert_eq!(out, vec!["Monday, January 1, 2024"]);
}

#[test]
fn message_format_number_default_style() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("x={0,number}", java.util.Locale.US); System.out.println(mf.format(new Object[]{1234.5}));"#,
    );
    assert_eq!(out, vec!["x=1,234.5"]);
}

#[test]
fn message_format_clone_independent_formatter() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("v={0}"); java.text.MessageFormat copy = (java.text.MessageFormat) mf.clone(); copy.applyPattern("copy={0}"); System.out.println(mf.format(new Object[]{"a"})); System.out.println(copy.format(new Object[]{"a"}));"#,
    );
    assert_eq!(out, vec!["v=a", "copy=a"]);
}

#[test]
fn message_format_equals_same_pattern() {
    let out = run_main(
        r#"java.text.MessageFormat a = new java.text.MessageFormat("p {0}"); java.text.MessageFormat b = new java.text.MessageFormat("p {0}"); System.out.println(a.equals(b));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn message_format_format_object_array_varargs_style() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0} / {1} / {2}"); System.out.println(mf.format(new Object[]{"a", "b", "c"}));"#,
    );
    assert_eq!(out, vec!["a / b / c"]);
}

#[test]
fn message_format_choice_with_fractional_boundary() {
    let out = run_main(
        r#"java.text.MessageFormat mf = new java.text.MessageFormat("{0,choice,0#none|0<some|2#two}"); System.out.println(mf.format(new Object[]{1.5}));"#,
    );
    assert_eq!(out, vec!["some"]);
}

#[test]
fn message_format_static_format_with_two_arguments() {
    let out = run_main(
        r#"System.out.println(java.text.MessageFormat.format("{0} wins by {1}", new Object[]{"Team A", 3}));"#,
    );
    assert_eq!(out, vec!["Team A wins by 3"]);
}
