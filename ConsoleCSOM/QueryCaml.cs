using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleCSOM
{
    public class QueryCaml
    {
        public static string QueryListItem = @"
        <View>
            <Query>
                <Where>
                    <Neq>
                        <FieldRef Name='Test_x0020_2'/>
                        <Value Type='Text'>about default</Value>
                    </Neq>
                </Where>
            </Query>
            <RowLimit>100</RowLimit>
        </View>";
        public static string QueryListView = @"
        <Where>
            <Eq>
                <FieldRef Name='City_x0020_Hunter'/>
                <Value Type='TaxonomyFieldTypeMulti'>Ho Chi Minh</Value>
            </Eq>
        </Where>
        <OrderBy><FieldRef Name='ID' Ascending='False'/></OrderBy>";
        public static string QueryUpdateListItem = @"
        <View>
            <Query>
                <Where>
                    <Neq>
                        <FieldRef Name='Test_x0020_2'/>
                        <Value Type='Text'>about default</Value>
                    </Neq>
                </Where>
            </Query>
            <RowLimit>100</RowLimit>
        </View>";
        

    }
}
