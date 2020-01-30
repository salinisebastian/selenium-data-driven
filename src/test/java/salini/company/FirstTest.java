package salini.company;

import java.io.IOException;
import java.util.ArrayList;

import org.testng.annotations.Test;

public class FirstTest {
	@Test

	public void test() throws IOException {
		// TODO Auto-generated method stub

		Data d = new Data();
		ArrayList<String> w = d.getData("Class C", "Sheet1");

		for (int i = 0; i < d.getnoOfColumns(); i++) {
			System.out.println(w.get(i));
		}

	}

}
