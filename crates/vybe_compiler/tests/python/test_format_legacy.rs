use crate::helpers::{run_print, run_python_one};

#[test]
fn percent_format_basic() {
    assert_eq!(run_print("'{}' % 'x'"), "x");
}

#[test]
fn percent_format_int() {
    assert_eq!(run_print("'%d' % 42"), "42");
}

#[test]
fn percent_format_multiple() {
    assert_eq!(run_print("'%s-%d' % ('a', 1)"), "a-1");
}

#[test]
fn percent_format_float() {
    assert_eq!(run_print("'%.2f' % 1.234"), "1.23");
}

#[test]
fn percent_format_width() {
    assert_eq!(run_print("'%5d' % 7"), "    7");
}

#[test]
fn percent_format_zero_pad() {
    assert_eq!(run_print("'%03d' % 4"), "004");
}

#[test]
fn percent_format_hex() {
    assert_eq!(run_print("'%x' % 255"), "ff");
}

#[test]
fn percent_format_octal() {
    assert_eq!(run_print("'%o' % 8"), "10");
}

#[test]
fn percent_format_dict_named() {
    assert_eq!(run_print("'%(name)s' % {'name': 'py'}"), "py");
}

#[test]
fn percent_format_dict_multiple_keys() {
    assert_eq!(run_print("'%(a)d%(b)d' % {'a': 1, 'b': 2}"), "12");
}

#[test]
fn format_method_positional() {
    assert_eq!(run_print("'{} {}'.format('a', 'b')"), "a b");
}

#[test]
fn format_method_indexed() {
    assert_eq!(run_print("'{1}-{0}'.format('b', 'a')"), "a-b");
}

#[test]
fn format_method_named() {
    assert_eq!(run_print("'{name}'.format(name='x')"), "x");
}

#[test]
fn format_method_width_align() {
    assert_eq!(run_print("'{:>3}'.format('x')"), "  x");
}

#[test]
fn format_method_precision_float() {
    assert_eq!(run_print("'{:.1f}'.format(2.25)"), "2.2");
}

#[test]
fn format_method_thousands_sep() {
    assert_eq!(run_print("'{:,}'.format(1000)"), "1,000");
}

#[test]
fn format_method_percent_style() {
    assert_eq!(run_print("'{:.0%}'.format(0.5)"), "50%");
}

#[test]
fn format_method_binary() {
    assert_eq!(run_print("'{:b}'.format(5)"), "101");
}

#[test]
fn format_method_hex() {
    assert_eq!(run_print("'{:x}'.format(255)"), "ff");
}

#[test]
fn format_method_escape_braces() {
    assert_eq!(run_print("'{{}}'.format()"), "{}");
}

#[test]
fn str_format_literal_percent_escape() {
    assert_eq!(run_print("'100%%' % ()"), "100%");
}

#[test]
fn percent_format_tuple_single_element() {
    assert_eq!(run_print("'%s' % ('a',)"), "a");
}

#[test]
fn percent_format_unicode_s() {
    assert_eq!(run_print("'%s' % 'hi'"), "hi");
}

#[test]
fn percent_format_repr_r() {
    assert_eq!(run_print("'%r' % 'a'"), "'a'");
}

#[test]
fn format_method_fill_char() {
    assert_eq!(run_print("'{:*>4}'.format('x')"), "***x");
}

#[test]
fn format_method_center() {
    assert_eq!(run_print("'{:^4}'.format('ab')"), " ab ");
}

#[test]
fn format_method_left() {
    assert_eq!(run_print("'{:<4}'.format('ab')"), "ab  ");
}

#[test]
fn format_method_sign_plus() {
    assert_eq!(run_print("'{:+d}'.format(3)"), "+3");
}

#[test]
fn format_method_sign_space() {
    assert_eq!(run_print("'{: d}'.format(3)"), " 3");
}

#[test]
fn format_method_scientific() {
    assert_eq!(run_print("'{:e}'.format(1000)"), "1.000000e+03");
}

#[test]
fn format_method_join_components() {
    assert_eq!(
        run_python_one("parts = ['a', 'b']\nprint('{0}-{1}'.format(*parts))\n"),
        "a-b"
    );
}

#[test]
fn format_method_unpack_mapping() {
    assert_eq!(run_print("'{a}-{b}'.format(**{'a': 1, 'b': 2})"), "1-2");
}

#[test]
fn percent_mapping_only() {
    assert_eq!(run_print("'%(x)s' % dict(x='y')"), "y");
}

#[test]
fn format_method_nested_field() {
    assert_eq!(run_print("'{0[1]}'.format([10, 20])"), "20");
}

#[test]
fn format_method_attr_access() {
    assert_eq!(
        run_python_one(
            "class P:\n def __init__(self):\n  self.x = 9\nprint('{0.x}'.format(P()))\n"
        ),
        "9"
    );
}

#[test]
fn percent_width_asterisk() {
    assert_eq!(run_print("'%*d' % (5, 7)"), "    7");
}

#[test]
fn percent_precision_asterisk() {
    assert_eq!(run_print("'%.*f' % (2, 1.2345)"), "1.23");
}

#[test]
fn format_method_zero_pad() {
    assert_eq!(run_print("'{:04d}'.format(7)"), "0007");
}

#[test]
fn format_method_negative_numbers() {
    assert_eq!(run_print("'{:d}'.format(-5)"), "-5");
}

#[test]
fn format_method_bool() {
    assert_eq!(run_print("'{!s}'.format(True)"), "True");
}

#[test]
fn format_method_repr_flag() {
    assert_eq!(run_print("'{!r}'.format('a')"), "'a'");
}

#[test]
fn format_method_none() {
    assert_eq!(run_print("'{!s}'.format(None)"), "None");
}

#[test]
fn percent_format_empty_tuple() {
    assert_eq!(run_print("'ok' % ()"), "ok");
}

#[test]
fn format_method_repeat_template() {
    assert_eq!(
        run_print("'-'.join(['{:02d}'.format(x) for x in range(3)])"),
        "00-01-02"
    );
}

#[test]
fn format_method_replace_after() {
    assert_eq!(
        run_python_one("s = '{:>3}'.format('x')\nprint(len(s))\n"),
        "3"
    );
}
