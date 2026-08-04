! vybe-test: fortran/derived_type_intrinsics_interop/test_derived_type_intrinsics_interop_bind_c_shape_query
! origin: languages/fortran/tests/fortran/test_derived_type_intrinsics_interop.rs

program test_derived_type_intrinsics_interop
    use iso_c_binding, only: c_int
    type, bind(C) :: payload
        integer(c_int) :: value
    end type

    type(payload) :: p
    p%value = 9
    if ((p%value) /= 9) then
    print *, "FAIL: want [9] got [", p%value, "]"
    stop 1
end if
end program test_derived_type_intrinsics_interop
