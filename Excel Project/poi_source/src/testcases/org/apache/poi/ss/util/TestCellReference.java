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

package org.apache.poi.ss.util;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellReference;

import org.junit.Test;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

/**
 * Tests that the common CellReference works as we need it to.
 * Note - some additional testing is also done in the HSSF class,
 *  {@link org.apache.poi.hssf.util.TestCellReference}
 */
public final class TestCellReference {
    @Test
    public void testConstructors() {
        CellReference cellReference;
        final String sheet = "Sheet1";
        final String cellRef = "A1";
        final int row = 0;
        final int col = 0;
        final boolean absRow = true;
        final boolean absCol = false;

        cellReference = new CellReference(row, col);
        assertEquals("A1", cellReference.formatAsString());

        cellReference = new CellReference(row, col, absRow, absCol);
        assertEquals("A$1", cellReference.formatAsString());

        cellReference = new CellReference(row, (short)col);
        assertEquals("A1", cellReference.formatAsString());

        cellReference = new CellReference(cellRef);
        assertEquals("A1", cellReference.formatAsString());

        cellReference = new CellReference(sheet, row, col, absRow, absCol);
        assertEquals("Sheet1!A$1", cellReference.formatAsString());
    }

    @Test
    public void testFormatAsString() {
        CellReference cellReference;

        cellReference = new CellReference(null, 0, 0, false, false);
        assertEquals("A1", cellReference.formatAsString());

        //absolute references
        cellReference = new CellReference(null, 0, 0, true, false);
        assertEquals("A$1", cellReference.formatAsString());

        //sheet name with no spaces
        cellReference = new CellReference("Sheet1", 0, 0, true, false);
        assertEquals("Sheet1!A$1", cellReference.formatAsString());

        //sheet name with spaces
        cellReference = new CellReference("Sheet 1", 0, 0, true, false);
        assertEquals("'Sheet 1'!A$1", cellReference.formatAsString());
    }

    @Test
    public void testGetCellRefParts() {
        CellReference cellReference;
        String[] parts;

        String cellRef = "A1";
        cellReference = new CellReference(cellRef);
        assertEquals(0, cellReference.getCol());
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("1", parts[1]);
        assertEquals("A", parts[2]);

        cellRef = "AA1";
        cellReference = new CellReference(cellRef);
        assertEquals(26, cellReference.getCol());
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("1", parts[1]);
        assertEquals("AA", parts[2]);

        cellRef = "AA100";
        cellReference = new CellReference(cellRef);
        assertEquals(26, cellReference.getCol());
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("100", parts[1]);
        assertEquals("AA", parts[2]);

        cellRef = "AAA300";
        cellReference = new CellReference(cellRef);
        assertEquals(702, cellReference.getCol());
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("300", parts[1]);
        assertEquals("AAA", parts[2]);

        cellRef = "ZZ100521";
        cellReference = new CellReference(cellRef);
        assertEquals(26*26+25, cellReference.getCol());
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("100521", parts[1]);
        assertEquals("ZZ", parts[2]);

        cellRef = "ZYX987";
        cellReference = new CellReference(cellRef);
        assertEquals(26*26*26 + 25*26 + 24 - 1, cellReference.getCol());
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("987", parts[1]);
        assertEquals("ZYX", parts[2]);

        cellRef = "AABC10065";
        cellReference = new CellReference(cellRef);
        parts = cellReference.getCellRefParts();
        assertNotNull(parts);
        assertEquals(null, parts[0]);
        assertEquals("10065", parts[1]);
        assertEquals("AABC", parts[2]);
    }

    @Test
    public void testGetColNumFromRef() {
        String cellRef = "A1";
        CellReference cellReference = new CellReference(cellRef);
        assertEquals(0, cellReference.getCol());

        cellRef = "AA1";
        cellReference = new CellReference(cellRef);
        assertEquals(26, cellReference.getCol());

        cellRef = "AB1";
        cellReference = new CellReference(cellRef);
        assertEquals(27, cellReference.getCol());

        cellRef = "BA1";
        cellReference = new CellReference(cellRef);
        assertEquals(26+26, cellReference.getCol());

        cellRef = "CA1";
        cellReference = new CellReference(cellRef);
        assertEquals(26+26+26, cellReference.getCol());

        cellRef = "ZA1";
        cellReference = new CellReference(cellRef);
        assertEquals(26*26, cellReference.getCol());

        cellRef = "ZZ1";
        cellReference = new CellReference(cellRef);
        assertEquals(26*26+25, cellReference.getCol());

        cellRef = "AAA1";
        cellReference = new CellReference(cellRef);
        assertEquals(26*26+26, cellReference.getCol());


        cellRef = "A1100";
        cellReference = new CellReference(cellRef);
        assertEquals(0, cellReference.getCol());

        cellRef = "BC15";
        cellReference = new CellReference(cellRef);
        assertEquals(54, cellReference.getCol());
    }

    @Test
    public void testGetRowNumFromRef() {
        String cellRef = "A1";
        CellReference cellReference = new CellReference(cellRef);
        assertEquals(0, cellReference.getRow());

        cellRef = "A12";
        cellReference = new CellReference(cellRef);
        assertEquals(11, cellReference.getRow());

        cellRef = "AS121";
        cellReference = new CellReference(cellRef);
        assertEquals(120, cellReference.getRow());
    }

    @Test
    public void testConvertNumToColString() {
        short col = 702;
        String collRef = new CellReference(0, col).formatAsString();
        assertEquals("AAA1", collRef);

        short col2 = 0;
        String collRef2 = new CellReference(0, col2).formatAsString();
        assertEquals("A1", collRef2);

        short col3 = 27;
        String collRef3 = new CellReference(0, col3).formatAsString();
        assertEquals("AB1", collRef3);

        short col4 = 2080;
        String collRef4 = new CellReference(0, col4).formatAsString();
        assertEquals("CBA1", collRef4);
    }

    @Test
    public void testBadRowNumber() {
        SpreadsheetVersion v97 = SpreadsheetVersion.EXCEL97;
        SpreadsheetVersion v2007 = SpreadsheetVersion.EXCEL2007;

        confirmCrInRange(true, "A", "1", v97);
        confirmCrInRange(true, "IV", "65536", v97);
        confirmCrInRange(false, "IV", "65537", v97);
        confirmCrInRange(false, "IW", "65536", v97);

        confirmCrInRange(true, "A", "1", v2007);
        confirmCrInRange(true, "XFD", "1048576", v2007);
        confirmCrInRange(false, "XFD", "1048577", v2007);
        confirmCrInRange(false, "XFE", "1048576", v2007);

        assertFalse("Identified bug 47312a", CellReference.cellReferenceIsWithinRange("B", "0", v97));

        confirmCrInRange(false, "A", "0", v97);
        confirmCrInRange(false, "A", "0", v2007);
    }

    @Test
    public void testInvalidReference() {
        try {
            new CellReference("Sheet1!#REF!");
            fail("Shouldn't be able to create a #REF! refence");
        } catch(IllegalArgumentException expected) {}

        try {
            new CellReference("'MySheetName'!#REF!");
            fail("Shouldn't be able to create a #REF! refence");
        } catch(IllegalArgumentException expected) {}

        try {
            new CellReference("#REF!");
            fail("Shouldn't be able to create a #REF! refence");
        } catch(IllegalArgumentException expected) {}
    }

    private static void confirmCrInRange(boolean expResult, String colStr, String rowStr,
            SpreadsheetVersion sv) {
        if (expResult == CellReference.cellReferenceIsWithinRange(colStr, rowStr, sv)) {
            return;
        }
        fail("expected (c='" + colStr + "', r='" + rowStr + "' to be "
                + (expResult ? "within" : "out of") + " bounds for version " + sv.name());
    }

    @Test
    public void testConvertColStringToIndex() {
        assertEquals(0, CellReference.convertColStringToIndex("A"));
        assertEquals(1, CellReference.convertColStringToIndex("B"));
        assertEquals(14, CellReference.convertColStringToIndex("O"));
        assertEquals(701, CellReference.convertColStringToIndex("ZZ"));
        assertEquals(18252, CellReference.convertColStringToIndex("ZZA"));

        assertEquals(0, CellReference.convertColStringToIndex("$A"));
        assertEquals(1, CellReference.convertColStringToIndex("$B"));

        try {
            CellReference.convertColStringToIndex("A$");
            fail("Should throw exception here");
        } catch (IllegalArgumentException expected) {
            assertTrue(expected.getMessage().contains("A$"));
        }
    }

    @Test
    public void testConvertNumColColString() {
        assertEquals("A", CellReference.convertNumToColString(0));
        assertEquals("AV", CellReference.convertNumToColString(47));
        assertEquals("AW", CellReference.convertNumToColString(48));
        assertEquals("BF", CellReference.convertNumToColString(57));

        assertEquals("", CellReference.convertNumToColString(-1));
        assertEquals("", CellReference.convertNumToColString(Integer.MIN_VALUE));
        assertEquals("", CellReference.convertNumToColString(Integer.MAX_VALUE));
        assertEquals("FXSHRXW", CellReference.convertNumToColString(Integer.MAX_VALUE-1));
    }
}
