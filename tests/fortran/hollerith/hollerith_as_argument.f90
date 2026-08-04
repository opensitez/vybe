! vybe-test: fortran/hollerith/hollerith_as_argument
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    call show(5Hhello)
contains
    subroutine show(msg)
        integer, intent(in) :: msg
        print *, 'received'
    end subroutine show
end program test
