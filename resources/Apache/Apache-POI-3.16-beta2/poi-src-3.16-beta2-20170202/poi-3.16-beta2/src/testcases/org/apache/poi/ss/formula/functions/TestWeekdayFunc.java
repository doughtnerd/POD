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

package org.apache.poi.ss.formula.functions;

import org.apache.poi.ss.formula.eval.*;
import org.junit.Test;

import static org.junit.Assert.*;


public class TestWeekdayFunc {
    @Test
    public void testEvaluate() throws Exception {
        assertEquals(2.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(2.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(1.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(1.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(2.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(0.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(3.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(1.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(11.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(7.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(12.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(6.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(13.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(5.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(14.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(4.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(15.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(3.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(16.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(2.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(17.0)}, 0, 0)).getNumberValue(), 0.001);

        assertEquals(3.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(3.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(1.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(2.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(2.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(1.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(3.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(2.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(11.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(1.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(12.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(7.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(13.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(6.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(14.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(5.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(15.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(4.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(16.0)}, 0, 0)).getNumberValue(), 0.001);
        assertEquals(3.0, ((NumberEval)WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(39448.0), new NumberEval(17.0)}, 0, 0)).getNumberValue(), 0.001);
    }

    @Test
    public void testEvaluateInvalid() throws Exception {
        assertEquals(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.evaluate(new ValueEval[]{}, 0, 0));
        assertEquals(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(1.0), new NumberEval(1.0)}, 0, 0));

        assertEquals(ErrorEval.NUM_ERROR, WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(-1.0)}, 0, 0));
        assertEquals(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.evaluate(new ValueEval[]{new StringEval("")}, 0, 0));
        assertEquals(ErrorEval.VALUE_INVALID, WeekdayFunc.instance.evaluate(new ValueEval[]{new StringEval("1"), new StringEval("")}, 0, 0));
        assertEquals(ErrorEval.NUM_ERROR, WeekdayFunc.instance.evaluate(new ValueEval[]{new StringEval("2"), BlankEval.instance}, 0, 0));
        assertEquals(ErrorEval.NUM_ERROR, WeekdayFunc.instance.evaluate(new ValueEval[]{new StringEval("3"), MissingArgEval.instance}, 0, 0));
        assertEquals(ErrorEval.NUM_ERROR, WeekdayFunc.instance.evaluate(new ValueEval[]{new NumberEval(1.0), new NumberEval(18.0)}, 0, 0));
    }
}