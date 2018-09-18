package com.xu.aboutpoi.toword;

import com.xu.aboutpoi.entity.Customer;
import com.xu.aboutpoi.service.toWord.CombineWord;
import com.xu.aboutpoi.service.toWord.PaseDataToWordTable;
import org.junit.BeforeClass;
import org.junit.Test;

import java.util.Arrays;
import java.util.List;

/**
 * @author xuhongda on 2018/9/17
 * com.xu.aboutpoi.toword
 * about-poi
 */







public class ToWordTest {
   /* Customer customer1 = new Customer();
    Customer customer2 = new Customer();
    Customer customer3 = new Customer();
    @BeforeClass
    public void be(){

        customer1.setContact("xxx");
        customer1.setEmail("xx@c.com");
        customer1.setId(1L);
        customer1.setName("cc1");
        customer1.setRemark("100|");
        customer1.setTelephone("131");

        customer2.setContact("xxx");
        customer2.setEmail("xx@c.com");
        customer2.setId(2L);
        customer2.setName("cc2");
        customer2.setRemark("200");
        customer2.setTelephone("131");

        customer3.setContact("xxx");
        customer3.setEmail("xx@c.com");
        customer3.setId(3L);
        customer3.setName("cc3");
        customer3.setRemark("300");
        customer3.setTelephone("131");
    }*/
    @Test
    public void test1() throws Exception {
        Customer customer1 = new Customer();
        Customer customer2 = new Customer();
        Customer customer3 = new Customer();
        customer1.setContact("xxx");
        customer1.setEmail("xx@c.com");
        customer1.setId(1L);
        customer1.setName("cc1");
        customer1.setRemark("100|");
        customer1.setTelephone("131");

        customer2.setContact("xxx");
        customer2.setEmail("xx@c.com");
        customer2.setId(2L);
        customer2.setName("cc2");
        customer2.setRemark("200");
        customer2.setTelephone("131");

        customer3.setContact("xxx");
        customer3.setEmail("xx@c.com");
        customer3.setId(3L);
        customer3.setName("cc3");
        customer3.setRemark("300");
        customer3.setTelephone("131");
        List<Customer> customers = Arrays.asList(customer1, customer2, customer3);
        PaseDataToWordTable.newWordDoc(Customer.class, customers, "customer",3);
    }

}
