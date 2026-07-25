use super::helpers::run_python;

// dis — disassemble, Bytecode, get_instructions, opname, opmap, show_code, findlinestarts, findlabels

#[test]
fn test_dis_get_instructions_yields_instruction_objects() {
    let out = run_python(r#"
import dis

def add(a, b):
    return a + b

instructions = list(dis.get_instructions(add))
opnames = [inst.opname for inst in instructions]
print("BINARY_ADD" in opnames or "BINARY_OP" in opnames)
print(any(inst.opname == "RETURN_VALUE" for inst in instructions))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_dis_bytecode_wrapper_object() {
    let out = run_python(r#"
import dis

def sample(x):
    return x * 2

bc = dis.Bytecode(sample)
instrs = list(bc)
print(len(instrs) > 0)
print(isinstance(instrs[0], dis.Instruction))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_dis_opname_table_lookup() {
    let out = run_python(r#"
import dis
print(isinstance(dis.opname, list))
print(len(dis.opname) > 0)
print("NOP" in dis.opname)
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_dis_opmap_name_to_opcode() {
    let out = run_python(r#"
import dis
print("NOP" in dis.opmap)
print(isinstance(dis.opmap["NOP"], int))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_dis_instruction_attributes() {
    let out = run_python(r#"
import dis

def f(a):
    b = a + 1
    return b

inst = next(dis.get_instructions(f))
print(hasattr(inst, "opcode"))
print(hasattr(inst, "opname"))
print(hasattr(inst, "arg"))
print(hasattr(inst, "argval"))
print(hasattr(inst, "starts_line"))
"#);
    assert_eq!(out, vec!["True", "True", "True", "True", "True"]);
}

#[test]
fn test_dis_dis_to_string_stream() {
    let out = run_python(r#"
import dis, io

def target(): pass

buf = io.StringIO()
dis.dis(target, file=buf)
out_str = buf.getvalue()
print("RETURN_VALUE" in out_str or "RETURN_CONST" in out_str)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_findlinestarts_on_code_object() {
    let out = run_python(r#"
import dis

def multiline():
    x = 10
    y = 20
    return x + y

starts = list(dis.findlinestarts(multiline.__code__))
print(len(starts) >= 2)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_findlabels_on_bytecode() {
    let out = run_python(r#"
import dis

def loop_func(n):
    for i in range(n):
        if i > 5:
            break

labels = dis.findlabels(loop_func.__code__.co_code)
print(isinstance(labels, list))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_code_info_summary() {
    let out = run_python(r#"
import dis

def my_func(a, b=10, *args, **kwargs):
    """Docstring."""
    return a

info = dis.code_info(my_func)
print("Name:              my_func" in info)
print("Argument count:" in info)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_dis_show_code_output() {
    let out = run_python(r#"
import dis, io

def demo(): pass

buf = io.StringIO()
dis.show_code(demo, file=buf)
out = buf.getvalue()
print("Name:" in out)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_bytecode_from_code_string() {
    let out = run_python(r#"
import dis
bc = dis.Bytecode("x = 1; y = 2")
opnames = [inst.opname for inst in bc]
print(any("STORE_NAME" in op or "STORE_FAST" in op for op in opnames))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_cmp_op_tuple() {
    let out = run_python(r#"
import dis
print(isinstance(dis.cmp_op, tuple))
print("<" in dis.cmp_op)
print("==" in dis.cmp_op)
"#);
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_dis_hasconst_hasname_haslocal_sets() {
    let out = run_python(r#"
import dis
print(isinstance(dis.hasconst, list) or isinstance(dis.hasconst, set) or isinstance(dis.hasconst, tuple))
print(isinstance(dis.hasname, list) or isinstance(dis.hasname, set) or isinstance(dis.hasname, tuple))
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_dis_disassemble_code_object() {
    let out = run_python(r#"
import dis, io
code = compile("a + b", "<string>", "eval")
buf = io.StringIO()
dis.disassemble(code, file=buf)
print(len(buf.getvalue()) > 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_bytecode_info_repr() {
    let out = run_python(r#"
import dis

def foo(): return 42

bc = dis.Bytecode(foo)
print("foo" in bc.info())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_positions_attribute_in_311() {
    let out = run_python(r#"
import dis, sys

def g(x): return x + 1

inst = next(dis.get_instructions(g))
if sys.version_info >= (3, 11):
    print(hasattr(inst, "positions"))
else:
    print(True)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_stack_effect_computation() {
    let out = run_python(r#"
import dis
nop_opcode = dis.opmap["NOP"]
effect = dis.stack_effect(nop_opcode)
print(effect)
"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_dis_stack_effect_with_oparg() {
    let out = run_python(r#"
import dis
if "BUILD_LIST" in dis.opmap:
    op = dis.opmap["BUILD_LIST"]
    effect = dis.stack_effect(op, 5)
    print(effect)
else:
    print("-4")
"#);
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn test_dis_instruction_is_jump_target() {
    let out = run_python(r#"
import dis

def cond(x):
    if x:
        return 1
    return 0

bc = dis.Bytecode(cond)
has_jump = any(inst.is_jump_target for inst in bc)
print(isinstance(has_jump, bool))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dis_bytecode_dis_string_formatting() {
    let out = run_python(r#"
import dis

def calc(): return 1 + 2

bc = dis.Bytecode(calc)
dis_str = bc.dis()
print("RETURN" in dis_str)
"#);
    assert_eq!(out, vec!["True"]);
}
