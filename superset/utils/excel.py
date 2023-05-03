# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
import io
from typing import Any
import pandas as pd


def dt_inplace(df: pd.DataFrame) -> pd.DataFrame:
    """Automatically detect and convert (in place!) each
    dataframe column of datatype 'object' to a datetime just
    when ALL of its non-NaN values can be successfully parsed
    by pd.to_datetime().  Also returns a ref. to df for
    convenient use in an expression.
    """
    from pandas.errors import ParserError
    for name in df.columns[df.dtypes == 'object']:
        try:
            df[name] = pd.to_datetime(df[name])
            df[name] = df[name].dt.tz_localize(None)
        except (ParserError, ValueError):
            pass
    return df


def df_to_excel(df: pd.DataFrame, **kwargs: Any) -> Any:
    df = dt_inplace(df)
    output = io.BytesIO()
    # pylint: disable=abstract-class-instantiated
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        return df.to_excel(writer, **kwargs)

