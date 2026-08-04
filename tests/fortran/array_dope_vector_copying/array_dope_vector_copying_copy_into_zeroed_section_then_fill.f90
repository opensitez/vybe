! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_into_zeroed_section_then_fill
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_into_zeroed_section_then_fill
    integer :: buffer(0:9)
    integer :: donor(2:6)
    donor = (/ 11, 22, 33, 44, 55 /)
    buffer(0:4) = 0
    buffer(2:6) = donor
    if ((sum(buffer)) /= 165) then
    print *, "FAIL: want [165] got [", sum(buffer), "]"
    stop 1
end if
    if ((buffer(0)) /= 0) then
    print *, "FAIL: want [0] got [", buffer(0), "]"
    stop 1
end if
    if ((buffer(6)) /= 55) then
    print *, "FAIL: want [55] got [", buffer(6), "]"
    stop 1
end if
    if ((buffer(2)) /= 11) then
    print *, "FAIL: want [11] got [", buffer(2), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_into_zeroed_section_then_fill
