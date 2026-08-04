! vybe-test: fortran/coarrays/co_reduce_user_op
! origin: languages/fortran/tests/fortran/test_coarrays.rs

program test
    integer :: x = this_image()
    call co_reduce(x, my_add, result_image=1)
    if (this_image() == 1) print *, x
contains
    pure function my_add(a, b) result(c)
        integer, intent(in) :: a, b
        integer :: c
        c = a + b
    end function my_add
end program test
