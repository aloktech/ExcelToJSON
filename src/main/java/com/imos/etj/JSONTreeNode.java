package com.imos.etj;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

/**
 *
 * @author p
 */
@Getter
@Setter
@ToString
public class JSONTreeNode {

    private final JSONData value;
    private final JSONData parent;
    private List<JSONTreeNode> children = Collections.synchronizedList(new ArrayList<>());

    public JSONTreeNode(JSONData data, JSONData parent) {
        this.value = data;
        this.parent = parent;
    }

}
