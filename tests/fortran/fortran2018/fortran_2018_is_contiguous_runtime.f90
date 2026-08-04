! vybe-test: fortran/fortran2018/fortran_2018_is_contiguous_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program t
    real, target :: a(10)
    real, pointer :: full(:)
    real, pointer :: stride(:)
    full => a
    stride => a(1:10:2)
    if ((is_contiguous(full)) /= 1) then
    print *, "FAIL: want [1] got [", is_contiguous(full), "]"
    stop 1
end if
    if ((is_contiguous(stride)) /= 0) then
    print *, "FAIL: want [0] got [", is_contiguous(stride), "]"
    stop 1
end if
end program t
