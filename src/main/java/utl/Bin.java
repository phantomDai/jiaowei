package utl;

import java.util.ArrayList;
import java.util.List;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/24
 */
public class Bin {
    List<String> list;

    public Bin() {
        this.list = new ArrayList<>();
    }

    public void add(String item){
        list.add(item);
    }

    public int size(){
        return list.size();
    }

    public String get(int index){
        return list.get(index);
    }

    public List<String> getList() {
        return list;
    }

    public void changeValue(int index, String updateValue){
        list.set(index, updateValue);
    }
}
