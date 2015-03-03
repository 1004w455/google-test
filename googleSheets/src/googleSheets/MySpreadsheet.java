package googleSheets;

import java.io.IOException;
import java.net.URL;
import java.util.List;

import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetEntry;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.util.ServiceException;

public class MySpreadsheet {

  public static void main(String[] args) throws IOException, ServiceException {

    // 먼저 서비스를 만들고 로그인 정보를 입력합니다.
    SpreadsheetService service = new SpreadsheetService("MySpreadsheetIntegration-v1");
    service.setUserCredentials("2014ksg@adwitt.com", "rkdtkdrb123");

    // Define the URL to request. This should never change.
    URL SPREADSHEET_FEED_URL = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");

    // feed.getEntries() 를 통해 등록되어 있는 모든 SpreadSheet 의 Entry 를 List 로 반환합니다.
    SpreadsheetFeed feed = service.getFeed(SPREADSHEET_FEED_URL, SpreadsheetFeed.class);

    List<SpreadsheetEntry> spreadsheets = feed.getEntries();

    // 전체적인 구조는 SpreadSheetEntry -> WorksheetEntry -> CellEntry 의 형태를 가지며, 동일한 방법에 의해 접근이 가능합니다.
    for (SpreadsheetEntry entry : spreadsheets) {
      System.out.println(entry.getTitle().getPlainText());

      List<WorksheetEntry> worksheets = entry.getWorksheets();
      for (WorksheetEntry worksheet : worksheets) {
        System.out.println("\t" + worksheet.getTitle().getPlainText());

        URL cellFeedUrl = worksheet.getCellFeedUrl();
        CellFeed cellFeed = service.getFeed(cellFeedUrl, CellFeed.class);
        for (CellEntry cell : cellFeed.getEntries()) {
          System.out.println(cell.getTitle().getPlainText());
          String shortId = cell.getId().substring(cell.getId().lastIndexOf('/') + 1);
          System.out.println(" -- Cell(" + shortId + "/" + cell.getTitle().getPlainText() + ") formula(" + cell.getCell().getInputValue() + ") numeric(" + cell.getCell().getNumericValue() + ") value(" + cell.getCell().getValue() + ")");
        }
      }
    }


  }

}
