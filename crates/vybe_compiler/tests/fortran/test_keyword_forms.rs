use super::helpers::compile_ok;

#[test]
fn fused_program_unit_keywords() {
    compile_ok(
        "module m\n  implicit none\n  interface\n    subroutine noop()\n    endsubroutine noop\n  endinterface\ncontains\n  integer function id(x) result(v)\n    integer, intent(in) :: x\n    integer :: v\n    v = x\n  endfunction id\nendmodule m\n",
    );
}