! vybe-test: fortran/hollerith/hollerith_in_common
! origin: languages/fortran/tests/fortran/test_hollerith.rs

program test
    integer :: label
    common /info/ label
    data label /4HINFO/
    print *, 'ok'
end program test
