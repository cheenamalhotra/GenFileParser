package test.java.com.cheenam.fileParser;

import java.util.List;

import org.junit.Before;
import org.junit.Test;

import junit.framework.TestCase;
import main.java.com.cheenam.fileParser.ExcelCSVParser;
import main.java.com.cheenam.fileParser.ExcelCSVParser.RowExtractor;

public class ExcelCSVParserTest extends TestCase {

	public static class Country {
		public String name;
		public String shortCode;

		public Country(String name, String shortCode) {
			this.shortCode = shortCode;
			this.name = name;
		}
	}

	interface Run {
		void run() throws Exception;
	}

	private ExcelCSVParser<Country> reader;

	private void checkList(List<Country> list) {
		assertEquals(list.get(1).shortCode, "AF");
		assertEquals(list.get(1).name, "AFGHANISTAN");
		assertEquals(list.get(57).shortCode, "CU");
		assertEquals(list.get(57).name, "CUBA");
		assertEquals(list.get(244).shortCode, "VI");
		assertEquals(list.get(244).name, "VIRGIN ISLANDS, U.S.");
	}

	public void delta(Run c) throws Exception {
		long start = System.currentTimeMillis();
		c.run();
		System.out.println("Delta: " + (System.currentTimeMillis() - start));
	}

	@Override
	@Before
	public void setUp() {
		RowExtractor<Country> converter = (row) -> new Country((String) row[0], (String) row[1]);
		reader = ExcelCSVParser.builder(Country.class).converter(converter).withHeader().csvDelimiter(',').sheets(1)
				.build();
	}

	@Test
	public void testBenchmark() throws Exception {
		RowExtractor<Country> converter = (row) -> new Country((String) row[0], (String) row[1]);
		ExcelCSVParser<Country> reader = ExcelCSVParser.builder(Country.class).converter(converter).csvDelimiter(',')
				.sheets(1).build();

		delta(() -> reader.read("src/test/resources/Countries.xlsx"));
		delta(() -> reader.read("src/test/resources/Countries.csv"));
	}

	@Test
	public void testShouldParseCorrectly_GivenCsvFile() throws Exception {
		List<Country> list;
		list = reader.read("src/test/resources/Countries.csv");
		checkList(list);
	}

	@Test
	public void testShouldParseCorrectly_GivenXlsxFile() throws Exception {
		List<Country> list;
		list = reader.read("src/test/resources/Countries.xlsx");
		checkList(list);
	}
}
