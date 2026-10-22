
export const qry_documentSearch: string = `
<View>
    <Query>
        <Where>            
            <And>
                <And>
                    <Geq>
                        <FieldRef Name="ID"/>
                        <Value Type="Number">{{minid}}</Value>
                    </Geq>
                    <Lt>
                        <FieldRef Name="ID"/>
                        <Value Type="Number">{{maxid}}</Value>
                    </Lt>
                </And>
                
                <And>
                    <And>
                        <And>
                            <And>
                                <{{Contains_EquipmentTypeSearchKey}}>
                                    <FieldRef Name="EquipmentType"/>
                                    <Value Type="Text"><![CDATA[{{EquipmentTypeSearchKey}}]]></Value>
                                </{{Contains_EquipmentTypeSearchKey}}>
                                <{{Contains_ManufactureSearchKey}}>
                                    <FieldRef Name="Manufacturer"/>
                                    <Value Type="Text"><![CDATA[{{ManufactureSearchKey}}]]></Value>
                                </{{Contains_ManufactureSearchKey}}>
                            </And>
                            
                            <And>
                                <{{Contains_ModelSearchKey}}>
                                    <FieldRef Name="Model"/>
                                    <Value Type="Text"><![CDATA[{{ModelSearchKey}}]]></Value>
                                </{{Contains_ModelSearchKey}}>
                                <{{Contains_AssetSearchKey}}>
                                    <FieldRef Name="Asset"/>
                                    <Value Type="Text"><![CDATA[{{AssetSearchKey}}]]></Value>
                                </{{Contains_AssetSearchKey}}>
                            </And>                                            
                        </And>                 

                        <And>
                            <And>
                                <{{Contains_DocumentClass_DiscplineSearchKey}}>
                                    <FieldRef Name="DocumentClass"/>
                                    <Value Type="Text"><![CDATA[{{DocumentClass_DiscplineSearchKey}}]]></Value>
                                </{{Contains_DocumentClass_DiscplineSearchKey}}>
                                <{{Contains_DocumentTypeSearchKey}}>
                                    <FieldRef Name="DocumentType"/>
                                    <Value Type="Text"><![CDATA[{{DocumentTypeSearchKey}}]]></Value>
                                </{{Contains_DocumentTypeSearchKey}}>
                            </And> 

                            <And>
                                <{{Contains_RegionSearchKey}}>
                                    <FieldRef Name="Region"/>
                                    <Value Type="Text"><![CDATA[{{RegionSearchKey}}]]></Value>
                                </{{Contains_RegionSearchKey}}>
                                <{{Contains_ClientSearchKey}}>
                                    <FieldRef Name="Client"/>
                                    <Value Type="Text"><![CDATA[{{ClientSearchKey}}]]></Value>
                                </{{Contains_ClientSearchKey}}>
                            </And>                         
                        </And>                        
                    </And>

                    <And>
                        <{{Contains_SerialSearchKey}}>
                            <FieldRef Name="Serial"/>
                            <Value Type="Text"><![CDATA[{{SerialSearchKey}}]]></Value>
                        </{{Contains_SerialSearchKey}}>

                        <And>
                            <{{Contains_SiteSearchKey}}>
                                <FieldRef Name="Site"/>
                                <Value Type="Text"><![CDATA[{{SiteSearchKey}}]]></Value>
                            </{{Contains_SiteSearchKey}}>
                            
                            <{{Contains_ProposalNumberSearchKey}}>
                                <FieldRef Name="ProposalNumber"/>
                                <Value Type="Text"><![CDATA[{{ProposalNumberSearchKey}}]]></Value>
                            </{{Contains_ProposalNumberSearchKey}}>                        
                        </And> 

                    </And>
                </And>    
            </And>            
        </Where>
    </Query>
</View>
`;

export const qry_myRequests: string = `
<View>
    <Query>
        <Where>
            <And>
                <And>
                    <Geq>
                        <FieldRef Name="ID"/>
                        <Value Type="Number">{{minid}}</Value>
                    </Geq>
                    <Lt>
                        <FieldRef Name="ID"/>
                        <Value Type="Number">{{maxid}}</Value>
                    </Lt>
                </And>
                <Eq>
                    <FieldRef Name="RequesterEmail"/>
                    <Value Type="Text">{{RequesterSearchKey}}</Value>
                </Eq>
            </And>
        </Where>
        <OrderBy>
            <FieldRef Name="ID" Ascending='FALSE'/>
        </OrderBy>
    </Query>
</View>
`;

export const qry_bizSegment: string = `
<View>
    <Query>
        <Where>
            <And>
                <And>
                    <Geq>
                        <FieldRef Name="ID"/>
                        <Value Type="Number">{{minid}}</Value>
                    </Geq>
                    <Lt>
                        <FieldRef Name="ID"/>
                        <Value Type="Number">{{maxid}}</Value>
                    </Lt>
                </And>
                <Eq>
                    <FieldRef Name="BizSegment"/>
                    <Value Type="Text">{{searchKey}}</Value>
                </Eq>
            </And>
        </Where>
        <OrderBy>
            <FieldRef Name="ID" Ascending='FALSE'/>
        </OrderBy>
    </Query>
</View>
`;
export const qry_segmentSearch: string = `
<View>
    <Query>
        <Where>            
            <Eq>
                <FieldRef Name="BizSegment"/>
                <Value Type="Choice">{{searchKey}}</Value>
            </Eq>
        </Where>
    </Query>
</View>
`;
export const qry_cifnoSearch: string = `
<View>
    <Query>
        <Where>            
            <Eq>
                <FieldRef Name="CIFNo"/>
                <Value Type="Text">{{cifno}}</Value>
            </Eq>
        </Where>
    </Query>
</View>
`;