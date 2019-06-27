package utl;

import java.io.File;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Set;

import static java.io.File.separator;

/**
 * describe:
 *
 * @author phantom
 * @date 2019/06/24
 */
public class GetOrganization {
    private String[] organizations;
    private String dir;

    public GetOrganization() {
        GetYears getYears = new GetYears();
        dir = getYears.getPath();
        parseDir();
    }

    private void parseDir(){
        File yearDirs = new File(dir);
        String[] yearFiles = yearDirs.list();
        Set<String> yearSet = new HashSet<>();
        for (String yearDir: yearFiles) {
            String tempOneYearDir = this.dir + separator + yearDir;
            File tempDir = new File(tempOneYearDir);
            String[] tempFileNames = tempDir.list();
            for (String tempName: tempFileNames) {
                yearSet.add(tempName);
            }
        }
        organizations = new String[1 + yearSet.size()];
        organizations[0] = "所有学校";
        String[] tempOrganizations = yearSet.toArray(new String[yearSet.size()]);
        System.arraycopy(tempOrganizations, 0, organizations,
                1, tempOrganizations.length);
    }

    public String[] getOrganizations() {
        return organizations;
    }
}
