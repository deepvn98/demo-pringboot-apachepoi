package com.javabase.apachepoidemo.repository;

import com.javabase.apachepoidemo.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;

import java.util.List;

public interface UserRepository extends JpaRepository<User, Long> {
    @Query(
            value = "select * from DEMO_APACEPOI.USER ", nativeQuery = true)
    List<User> findAllUser();


}
