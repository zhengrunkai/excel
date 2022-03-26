package com.example.excel.dao;

import com.example.excel.bean.Person;
import org.apache.ibatis.annotations.Mapper;

import java.util.List;

@Mapper
public interface PersonDao {

    List<Person> getPersonList();

    void insertPersonByBatch(List<Person> list);

}
