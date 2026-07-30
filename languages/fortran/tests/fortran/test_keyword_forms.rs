use super::helpers::compile_ok;

#[test]
fn fused_program_unit_keywords() {
    compile_ok(
        r#"
module m
    implicit none
    public

    interface
        subroutine noop()
        end subroutine noop
    end interface

contains
    pure function id(x) result(v)
        integer, intent(in) :: x
        integer :: v
        v = x
    end function id
end module m
"#,
    );
}

#[test]
fn block_construct_keyword_endings() {
    compile_ok(
        r#"
program kw_blocks
    implicit none
    integer :: a(1:3)
    integer :: i

    do i = 1, 3
        a(i) = i
    end do

    do i = 1, 3
        if (a(i) == 2) then
            a(i) = a(i) + 1
        else if (a(i) == 3) then
            a(i) = a(i) - 1
        else
            a(i) = a(i)
        end if
end do

    select case (a(2))
    case (1:2)
        a(2) = a(2) * 2
    case default
        a(2) = 0
    end select

    where (a > 0)
        a = a + 1
    end where
end program kw_blocks
"#,
    );
}

#[test]
fn procedure_keyword_forms() {
    compile_ok(
        r#"
module kw_proc
    implicit none
contains
    recursive integer function fact(n) result(r)
        integer, intent(in) :: n
        if (n <= 1) then
            r = 1
        else
            r = n * fact(n - 1)
        end if
    end function fact

    pure elemental integer function inc(x) result(r)
        integer, intent(in) :: x
        r = x + 1
    end function inc

    module subroutine noop_sub(x)
        integer, intent(inout) :: x
        x = x
    end subroutine noop_sub
end module kw_proc
"#,
    );
}

#[test]
fn interface_and_binding_keyword_forms() {
    compile_ok(
        r#"
module kw_bind_mod
    use, intrinsic :: iso_c_binding, only: c_int
    implicit none
    contains

subroutine kw_binding(x) bind(c, name="kw_binding")
    integer(c_int), intent(inout) :: x
    x = x
end subroutine kw_binding

subroutine call_kw_binding()
    integer(c_int) :: v
    call kw_binding(v)
end subroutine call_kw_binding
end module kw_bind_mod
"#,
    );
}

#[test]
fn abstract_interface_keyword_form() {
    compile_ok(
        r#"
module kw_abstract_iface
interface
    subroutine abstract_target(x)
        integer, intent(inout) :: x
    end subroutine abstract_target
end interface

abstract interface
    subroutine abstract_target(x)
        integer, intent(inout) :: x
    end subroutine abstract_target
end interface
end module kw_abstract_iface
"#,
    );
}
