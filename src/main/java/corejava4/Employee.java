package corejava4;

import lombok.Getter;
import lombok.Setter;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Getter
@Setter
public class Employee {
    private String id;
    private String name;
    private double totalSalaryFromFile;
    private double calculatedTotalSalary;
    private Map<String, Double> shiftsRate;
    private List<WorkingDay> workingDays;

    public Employee(String id, String name, double totalSalaryFromFile) {
        this.id = id;
        this.name = name;
        this.totalSalaryFromFile = totalSalaryFromFile;
        this.calculatedTotalSalary = 0;
        this.shiftsRate = new LinkedHashMap<>();
        this.workingDays = new ArrayList<>();
    }

    public void addWorkingDay(WorkingDay workingDay) {
        workingDays.add(workingDay);
        calculatedTotalSalary += workingDay.getTotalAmount();
    }

    @Override
    public String toString() {
        return "Employee{" +
                "id='" + id + '\'' +
                ", name='" + name + '\'' +
                ", totalSalaryFromFile=" + totalSalaryFromFile +
                ", calculatedTotalSalary=" + calculatedTotalSalary +
                ", workingDays=" + workingDays +
                '}';
    }
}
