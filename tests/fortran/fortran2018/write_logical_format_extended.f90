! vybe-test: fortran/fortran2018/write_logical_format_extended
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    logical :: flags(3) = [.true., .false., .true.]
    write(*, '(3L5)') flags
end program test
