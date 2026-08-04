! vybe-test: fortran/module_use_extended/compile_interface_operator_equality
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module eq_iface
    implicit none
    type :: Tag
        integer :: id
    end type Tag
    interface operator(==)
        module procedure tags_equal
    end interface
contains
    function tags_equal(a, b) result(same)
        type(Tag), intent(in) :: a, b
        logical :: same
        same = a%id == b%id
    end function tags_equal
end module eq_iface

program t
    use eq_iface
    type(Tag) :: x, y
    x%id = 1
    y%id = 1
    print *, x == y
end program t
