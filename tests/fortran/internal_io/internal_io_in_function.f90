! vybe-test: fortran/internal_io/internal_io_in_function
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: s
    s = int_to_str(42)
    print *, trim(s)
contains
    function int_to_str(n) result(s)
        integer, intent(in) :: n
        character(len=20) :: s
        write(s, '(I0)') n
    end function int_to_str
end program test
