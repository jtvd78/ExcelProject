/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.hssf.extractor;

import static org.apache.poi.POITestCase.assertContains;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;

import org.apache.poi.EmptyFileException;
import org.apache.poi.POIDataSamples;
import org.apache.poi.hssf.HSSFTestDataSamples;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.junit.Ignore;
import org.junit.Test;

/**
 * Unit tests for the Excel 5/95 and Excel 4 (and older) text 
 *  extractor
 */
public final class TestOldExcelExtractor {
    private static OldExcelExtractor createExtractor(String sampleFileName) throws IOException {
        File file = HSSFTestDataSamples.getSampleFile(sampleFileName);
        return new OldExcelExtractor(file);
    }

    @Test
    public void testSimpleExcel3() throws Exception {
        OldExcelExtractor extractor = createExtractor("testEXCEL_3.xls");

        // Check we can call getText without error
        String text = extractor.getText();

        // Check we find a few words we expect in there
        assertContains(text, "Season beginning August");
        assertContains(text, "USDA");
        
        // Check we find a few numbers we expect in there
        assertContains(text, "347");
        assertContains(text, "228");
        
        // Check we find a few string-literal dates in there
        assertContains(text, "1981/82");
        
        // Check the type
        assertEquals(3, extractor.getBiffVersion());
        assertEquals(0x10, extractor.getFileType());

        extractor.close();
    }
    

    @Test
    public void testSimpleExcel3NoReading() throws Exception {
        OldExcelExtractor extractor = createExtractor("testEXCEL_3.xls");
        assertNotNull(extractor);

        extractor.close();
    }

    @Test
    public void testSimpleExcel4() throws Exception {
        OldExcelExtractor extractor = createExtractor("testEXCEL_4.xls");

        // Check we can call getText without error
        String text = extractor.getText();

        // Check we find a few words we expect in there
        assertContains(text, "Size");
        assertContains(text, "Returns");
        
        // Check we find a few numbers we expect in there
        assertContains(text, "11");
        assertContains(text, "784");
        
        // Check the type
        assertEquals(4, extractor.getBiffVersion());
        assertEquals(0x10, extractor.getFileType());

        extractor.close();
    }
    
    @Test
    public void testSimpleExcel5() throws Exception {
        for (String ver : new String[] {"5", "95"}) {
            OldExcelExtractor extractor = createExtractor("testEXCEL_"+ver+".xls");
    
            // Check we can call getText without error
            String text = extractor.getText();
    
            // Check we find a few words we expect in there
            assertContains(text, "Sample Excel");
            assertContains(text, "Written and saved");
            
            // Check we find a few numbers we expect in there
            assertContains(text, "15");
            assertContains(text, "169");
            
            // Check we got the sheet names (new formats only)
            assertContains(text, "Sheet: Feuil3");
            
            // Check the type
            assertEquals(5, extractor.getBiffVersion());
            assertEquals(0x05, extractor.getFileType());

            extractor.close();
        }
    }

    @Test
    public void testStrings() throws Exception {
        OldExcelExtractor extractor = createExtractor("testEXCEL_4.xls");
        String text = extractor.getText();

        // Simple strings
        assertContains(text, "Table 10 -- Examination Coverage:");
        assertContains(text, "Recommended and Average Recommended Additional Tax After");
        assertContains(text, "Individual income tax returns, total");
        
        // More complicated strings
        assertContains(text, "$100,000 or more");
        assertContains(text, "S corporation returns, Form 1120S [10,15]");
        assertContains(text, "individual income tax return \u201Cshort forms.\u201D");
        
        // Formula based strings
        // TODO Find some then test

        extractor.close();
    }

    @Test
    public void testFormattedNumbersExcel4() throws Exception {
        OldExcelExtractor extractor = createExtractor("testEXCEL_4.xls");
        String text = extractor.getText();

        // Simple numbers
        assertContains(text, "151");
        assertContains(text, "784");
        
        // Numbers which come from formulas
        assertContains(text, "0.398"); // TODO Rounding
        assertContains(text, "624");
        
        // Formatted numbers
        // TODO
//      assertContains(text, "55,624");
//      assertContains(text, "11,743,477");

        extractor.close();
    }
    
    @Test
    public void testFormattedNumbersExcel5() throws Exception {
        for (String ver : new String[] {"5", "95"}) {
            OldExcelExtractor extractor = createExtractor("testEXCEL_"+ver+".xls");
            String text = extractor.getText();
            
            // Simple numbers
            assertContains(text, "1");
            
            // Numbers which come from formulas
            assertContains(text, "13");
            assertContains(text, "169");
            
            // Formatted numbers
            // TODO
//          assertContains(text, "100.00%");
//          assertContains(text, "155.00%");
//          assertContains(text, "1,125");
//          assertContains(text, "189,945");
//          assertContains(text, "1,234,500");
//          assertContains(text, "$169.00");
//          assertContains(text, "$1,253.82");

            extractor.close();
        }
    }
    
    @Test
    public void testFromFile() throws Exception {
        for (String ver : new String[] {"4", "5", "95"}) {
            String filename = "testEXCEL_"+ver+".xls";
            File f = HSSFTestDataSamples.getSampleFile(filename);
            
            OldExcelExtractor extractor = new OldExcelExtractor(f);
            String text = extractor.getText();
            assertNotNull(text);
            assertTrue(text.length() > 100);

            extractor.close();
        }
    }

    @Test
    public void testOpenInvalidFile() throws Exception {
        // a file that exists, but is a different format
        try {
            createExtractor("WithVariousData.xlsx");
            fail("Should catch Exception here");
        } catch (OfficeXmlFileException e) {
            // expected here
        }
        
    }

    @Test
    public void testOpenNonExistingFile() throws Exception {
        // a file that exists, but is a different format
        try {
            OldExcelExtractor extractor = new OldExcelExtractor(new File("notexistingfile.xls"));
            extractor.close();
            fail("Should catch Exception here");
        } catch (EmptyFileException e) {
            // expected here
        }
    }
    
    @Test
    public void testInputStream() throws Exception {
        File file = HSSFTestDataSamples.getSampleFile("testEXCEL_3.xls");
        InputStream stream = new FileInputStream(file);
        try {
            OldExcelExtractor extractor = new OldExcelExtractor(stream);
            String text = extractor.getText();
            assertNotNull(text);
            extractor.close();
        } finally {
            stream.close();
        }
    }

    @Test
    public void testInputStreamNPOIHeader() throws Exception {
        File file = HSSFTestDataSamples.getSampleFile("FormulaRefs.xls");
        InputStream stream = new FileInputStream(file);
        try {
            OldExcelExtractor extractor = new OldExcelExtractor(stream);
            extractor.close();
        } finally {
            stream.close();
        }
    }

    @Test
    public void testNPOIFSFileSystem() throws Exception {
        File file = HSSFTestDataSamples.getSampleFile("FormulaRefs.xls");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
        try {
            OldExcelExtractor extractor = new OldExcelExtractor(fs);
            extractor.close();
        } finally {
            fs.close();
        }
    }

    @Test
    public void testDirectoryNode() throws Exception {
        File file = HSSFTestDataSamples.getSampleFile("FormulaRefs.xls");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
        try {
            OldExcelExtractor extractor = new OldExcelExtractor(fs.getRoot());
            extractor.close();
        } finally {
            fs.close();
        }
    }

    @Test
    public void testDirectoryNodeInvalidFile() throws Exception {
        File file = POIDataSamples.getDocumentInstance().getFile("test.doc");
        NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
        try {
            OldExcelExtractor extractor = new OldExcelExtractor(fs.getRoot());
            extractor.close();
            fail("Should catch exception here");
        } catch (FileNotFoundException e) {
            // expected here
        } finally {
            fs.close();
        }
    }

    @Ignore("Calls System.exit()")
    @Test
    public void testMainUsage() throws Exception {
        PrintStream save = System.err;
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try {
                PrintStream str = new PrintStream(out, false, "UTF-8");
                System.setErr(str);
                OldExcelExtractor.main(new String[] {});
            } finally {
                out.close();
            }
        } finally {
            System.setErr(save);
        }
    }

    @Test
    public void testMain() throws Exception {
        File file = HSSFTestDataSamples.getSampleFile("testEXCEL_3.xls");
        PrintStream save = System.out;
        try {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try {
                PrintStream str = new PrintStream(out, false, "UTF-8");
                System.setOut(str);
                OldExcelExtractor.main(new String[] {file.getAbsolutePath()});
            } finally {
                out.close();
            }
            String string = new String(out.toByteArray(), "UTF-8");
            assertTrue("Had: " + string, 
                    string.contains("Table C-13--Lemons"));
        } finally {
            System.setOut(save);
        }
    }
}
