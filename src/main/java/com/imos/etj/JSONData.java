/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.imos.etj;

import lombok.AllArgsConstructor;
import lombok.EqualsAndHashCode;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 *
 * @author p
 */
@Getter
@Setter
@EqualsAndHashCode(exclude = {"data"})
@ToString
@AllArgsConstructor
public class JSONData {

    private JSONDataType dataType;
    private JSONValueType valueType;
    private Object data;
    private String key;
    private String parentKey;
    private JSONDataType parentDataType;
}
