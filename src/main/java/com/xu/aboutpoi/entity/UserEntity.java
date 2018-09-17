package com.xu.aboutpoi.entity;



import com.xu.aboutpoi.enums.UserSexEnum;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.ToString;

import java.io.Serializable;

@Data
@AllArgsConstructor
@NoArgsConstructor
@ToString
public class UserEntity implements Serializable {

    private static final long serialVersionUID = 1L;

    private Long id;
    private String userName;
    private String passWord;
    private UserSexEnum userSex;
    private String nickName;

}