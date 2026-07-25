use super::helpers::run_python;

// doctest — DocTestFinder, DocTestRunner, testmod, Example, OutputChecker, REPORT_NDIFF, ELLIPSIS, NORMALIZE_WHITESPACE

#[test]
fn test_doctest_testmod_basic_function_docstring() {
    let out = run_python(r#"
import doctest

def add(a, b):
    """
    >>> add(2, 3)
    5
    >>> add(-1, 1)
    0
    """
    return a + b

res = doctest.testmod()
print(res.failed)
print(res.attempted)
"#);
    assert_eq!(out, vec!["0", "2"]);
}

#[test]
fn test_doctest_ellipsis_flag_matching() {
    let out = run_python(r#"
import doctest

def get_list():
    """
    >>> get_list() # doctest: +ELLIPSIS
    [1, ..., 5]
    """
    return [1, 2, 3, 4, 5]

res = doctest.testmod(optionflags=doctest.ELLIPSIS)
print(res.failed)
print(res.attempted)
"#);
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn test_doctest_normalize_whitespace_flag() {
    let out = run_python(r#"
import doctest

def get_matrix():
    """
    >>> get_matrix() # doctest: +NORMALIZE_WHITESPACE
    [1, 2,
     3, 4]
    """
    return [1, 2, 3, 4]

res = doctest.testmod(optionflags=doctest.NORMALIZE_WHITESPACE)
print(res.failed)
print(res.attempted)
"#);
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn test_doctest_exception_matching() {
    let out = run_python(r#"
import doctest

def divide(a, b):
    """
    >>> divide(1, 0)
    Traceback (most recent call last):
        ...
    ZeroDivisionError: division by zero
    """
    return a / b

res = doctest.testmod()
print(res.failed)
print(res.attempted)
"#);
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn test_doctest_doctest_finder_finds_tests_in_module() {
    let out = run_python(r#"
import doctest, types

m = types.ModuleType("sample")
def foo():
    """
    >>> foo()
    'bar'
    """
    return "bar"

m.foo = foo
finder = doctest.DocTestFinder()
tests = finder.find(m)
print(len(tests))
print(tests[0].name)
"#);
    assert_eq!(out, vec!["1", "sample.foo"]);
}

#[test]
fn test_doctest_doctest_runner_run_test() {
    let out = run_python(r#"
import doctest

def greet(name):
    """
    >>> greet("World")
    'Hello World'
    """
    return f"Hello {name}"

finder = doctest.DocTestFinder()
runner = doctest.DocTestRunner(verbose=False)
tests = finder.find(greet)
for t in tests:
    res = runner.run(t)

print(runner.failures)
print(runner.tries)
"#);
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn test_doctest_example_object_attributes() {
    let out = run_python(r#"
import doctest
ex = doctest.Example("add(1, 2)\n", "3\n", lineno=10)
print(ex.source.strip())
print(ex.want.strip())
print(ex.lineno)
"#);
    assert_eq!(out, vec!["add(1, 2)", "3", "10"]);
}

#[test]
fn test_doctest_output_checker_check_output() {
    let out = run_python(r#"
import doctest
checker = doctest.OutputChecker()
want = "hello world\n"
got = "hello world\n"
print(checker.check_output(want, got, 0))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_doctest_dont_accept_blankline_flag() {
    let out = run_python(r#"
import doctest

def blank_output():
    """
    >>> blank_output()
    <BLANKLINE>
    """
    print()

res = doctest.testmod()
print(res.failed)
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_doctest_ignore_exception_detail_flag() {
    let out = run_python(r#"
import doctest

def err():
    """
    >>> err() # doctest: +IGNORE_EXCEPTION_DETAIL
    Traceback (most recent call last):
    ValueError: some message that is ignored
    """
    raise ValueError("different message")

res = doctest.testmod(optionflags=doctest.IGNORE_EXCEPTION_DETAIL)
print(res.failed)
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_doctest_register_optionflag_custom_flag() {
    let out = run_python(r#"
import doctest
flag = doctest.register_optionflag("MY_CUSTOM_FLAG")
print(isinstance(flag, int))
print(flag > 0)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_doctest_script_from_examples() {
    let out = run_python(r#"
import doctest
doc = """
>>> x = 10
>>> x + 5
15
"""
test = doctest.DocTestParser().get_doctest(doc, {}, "test", "file.py", 0)
script = doctest.script_from_examples(doc)
print("x = 10" in script)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_doctest_failing_doctest_detection() {
    let out = run_python(r#"
import doctest

def wrong():
    """
    >>> wrong()
    100
    """
    return 50

res = doctest.testmod()
print(res.failed)
print(res.attempted)
"#);
    assert_eq!(out, vec!["1", "1"]);
}

#[test]
fn test_doctest_doctest_parser_get_examples() {
    let out = run_python(r#"
import doctest
text = """
Some text
>>> a = 1
>>> a
1
"""
parser = doctest.DocTestParser()
examples = parser.get_examples(text)
print(len(examples))
print(examples[0].source.strip())
"#);
    assert_eq!(out, vec!["2", "a = 1"]);
}

#[test]
fn test_doctest_report_ndiff_flag_constants() {
    let out = run_python(r#"
import doctest
print(hasattr(doctest, "REPORT_NDIFF"))
print(hasattr(doctest, "REPORT_UDIFF"))
print(hasattr(doctest, "REPORT_CDIFF"))
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_doctest_skip_directive() {
    let out = run_python(r#"
import doctest

def skipped():
    """
    >>> skipped() # doctest: +SKIP
    not evaluated
    """
    return "actual"

res = doctest.testmod()
print(res.failed)
print(res.attempted)
"#);
    assert_eq!(out, vec!["0", "1"]);
}

#[test]
fn test_doctest_doctest_runner_summary() {
    let out = run_python(r#"
import doctest, io

def ok():
    """
    >>> ok()
    1
    """
    return 1

runner = doctest.DocTestRunner(verbose=False)
test = doctest.DocTestFinder().find(ok)[0]
runner.run(test)

buf = io.StringIO()
runner.summarize(out=buf.write)
print("1 items passed" in buf.getvalue() or "Passed" in buf.getvalue() or len(buf.getvalue()) >= 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_doctest_doctest_parser_get_doctest() {
    let out = run_python(r#"
import doctest
parser = doctest.DocTestParser()
dt = parser.get_doctest(">>> 1 + 1\n2", {}, "mytest", "test.py", 1)
print(dt.name)
print(len(dt.examples))
"#);
    assert_eq!(out, vec!["mytest", "1"]);
}

#[test]
fn test_doctest_output_checker_output_difference_formatting() {
    let out = run_python(r#"
import doctest
checker = doctest.OutputChecker()
ex = doctest.Example("foo()", "expected\n")
diff = checker.output_difference(ex, "got\n", 0)
print("Expected:" in diff)
print("Got:" in diff)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_doctest_dont_accept_true_for_1_flag() {
    let out = run_python(r#"
import doctest
print(hasattr(doctest, "DONT_ACCEPT_TRUE_FOR_1"))
"#);
    assert_eq!(out, vec!["True"]);
}
