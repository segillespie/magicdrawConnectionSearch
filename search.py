def search(element, previous, df_part, df_connector, df_port):

    print('Searching: ' + element)

    import os
    os.chdir("c:/Users/stephen.gillespie/Desktop/Mac/")
    from search import search
    from support_functions import all_connectors_attached_to_element
    from support_functions import port_on_part
    from support_functions import all_connectors_attached_to_port
    from support_functions import port_to_connectors
    from support_functions import all_connectors_owned_by_element
    from support_functions import element_idx
    from support_functions import ref_to_part
    from support_functions import part_between
    

################################################################################
    previous.append(element)
    CI_Set = ['CI', 'Ci', 'ci', 'cI', 'Y', 'y', 'Yes', 'yes', 'YES', 'C', 'c', 'Sub-CI']
    Final_Result = []
################################################################################    
    # bring in relevant info: e_type, e_idx
        # check to see what file element is located in; if element cannot be found, return end of path
    if element in df_connector['Element ID'].values:
        e_type = 'connector'
        e_idx = df_connector[df_connector['Element ID'] == element].index.tolist()
        isabstract = df_connector['Is Abstract'][e_idx[0]]
        isSoftware = False
        isWiringHarness = False
    elif element in df_part['Elm ID'].values:
        e_type = 'part'
        e_idx = df_part[df_part['Elm ID'] == element].index.tolist()
        isabstract = df_part['Is Abstract'][e_idx[0]]
        isSoftware = df_part['Is Software'][e_idx[0]]
        isWiringHarness = df_part['Is Wire Harness'][e_idx[0]]
    elif element in df_port['Elm ID'].values:
        e_type = 'port'
        e_idx = df_port[df_port['Elm ID'] == element].index.tolist()
        isabstract = df_port['Is Abstract'][e_idx[0]]
        isSoftware = False
        isWiringHarness = False
    else:
        isabstract = False
        isSoftware = False
        isWiringHarness = False
        del previous[len(previous) - 1]
        return [['End of Path: Element not in Ports, Parts, or Connectors DF']]


################################################################################
    # All possible situations for a given element I'm searching    

################################################################################
    # Condition 0: element is labeld as abstract / software / wiring harness in the model and therefore we do not want to continue searching this path.
    if isabstract:
        del previous[len(previous) - 1]
        return [[element, 'Last element abstract, search terminated']]
    elif isSoftware:
        del previous[len(previous) - 1]
        return [[element, 'Last element software, search terminated']]
    elif isWiringHarness:
        del previous[len(previous) - 1]
        return [[element, 'Last element wiring harness, search terminated']]
  
################################################################################
    # Condition 1: Element is a part property & First element searched
    elif e_type == 'part' and len(previous) <= 1:
        # then search every connector
        connectors = all_connectors_attached_to_element(element, df_connector, df_part, df_port)
        if connectors == []:
            del previous[len(previous) - 1]
            return [[element, 'End of Path: Part has no connectors']]
        for c in connectors:
            result = search(c, previous, df_part, df_connector, df_port)
            for r in result:
                r.insert(0, element)
                Final_Result.append(r)
        del previous[len(previous) - 1]
        return Final_Result
            
################################################################################            
    # Condition 2: Element is a part property & a CI
    elif e_type == 'part' and df_part['CI Indicator (System Context)'][e_idx[0]] in CI_Set:
        del previous[len(previous) - 1]
        return [[element, 'End Path, Last Element is CI']]
        
################################################################################            
    # Condition 3: Element is a part property & not a CI (i.e. either an assembly or a piece of a CI)
    elif e_type == 'part' and (not (df_part['CI Indicator (System Context)'][e_idx[0]] in CI_Set)):
        # look at all connectors that 1) are connected to part property and 2) share a port with the last connector and 3) haven't been previously searched
        print('Condition 3')
        # get all connectors that are attached to the current part
        connectors_to_part = all_connectors_attached_to_element(element, df_connector, df_part, df_port) 
        connectors_in_part = all_connectors_owned_by_element(df_part['Type'][e_idx[0]], df_connector)
        
        other_connectors = []
        for c in df_part['Connected Parts or Ports by Elm ID_split'][e_idx[0]]:
            if c in df_connector['Element ID'].values:
                other_connectors.append(c)
        
        # get previous connector index
        previous_connector = previous[len(previous) - 2] # the last previous element is the part itself, the element before that is the connector that got us to this element
        c_idx = df_connector[df_connector['Element ID'] == previous_connector].index.tolist() #get the index of the previous connector
        
        #find the current role that we're looking at
        current_role = port_on_part(c_idx, e_idx, df_connector, df_part, df_port)
        
        #get all connectors attached to this role
        connectors_to_port = all_connectors_attached_to_port(current_role[0], df_connector)[0]
        
        #take the intersection of connectors to the current part and to the current role
        connectors = list(set(connectors_to_part + connectors_in_part + other_connectors).intersection(connectors_to_port))
        
        # delete any connectors in the set that have previously been explored
        for p in previous: 
            if p in connectors: connectors.remove(p)
            
        # if there are no connectors to explore, this is the end of the path
        if connectors == []:
            del previous[len(previous) - 1]
            return [[element, 'End of Path: No connectors from current port and part to another. Not a CI']]
        
        # if there are connectors to explore, follow the methodology for exploring those connectors
        for c in connectors:
            result = search(c, previous, df_part, df_connector, df_port)
            for r in result:
                r.insert(0, element)
                Final_Result.append(r)
            del previous[len(previous) - 1]
            return Final_Result



#########################################################################################################################################################################################################
    # Case 4: the element is a connector.  New search version.
    elif e_type == 'connector':
        #added to stop double searching where connectors are display only in the model
        if type(df_connector['Name'][e_idx[0]]) == str:
            if 'For Display' in df_connector['Name'][e_idx[0]]:
                [[element, 'Error: Connector is for display purposes only.']]
        
        last_role = []
        previous_ID = previous[len(previous) - 2]
        ############################################################################
        # find the index and type of the previous element
        if previous_ID in df_connector['Element ID'].values:
            previous_type = 'connector'
            previous_idx = df_connector[df_connector['Element ID'] == previous_ID].index.tolist()
        elif previous_ID in df_part['Elm ID'].values:
            previous_type = 'part'
            previous_idx = df_part[df_part['Elm ID'] == previous_ID].index.tolist()
        elif previous_ID in df_port['Elm ID'].values:
            previous_type = 'port'
            previous_idx = df_port[df_port['Elm ID'] == previous_ID].index.tolist()
        else:
            del previous[len(previous) - 1]
            return [[element, 'Error: Previously searched element not in database.']]    
        ############################################################################    
        
        connector_roles = list(df_connector['ElmID of Role of Connector Ends (Port)_split'][e_idx[0]])
        
        # Ensure there are roles to look at.  if this is empty, end the search.
        if not connector_roles:
            del previous[len(previous) - 1]
            return [[element, 'Error, connector has no roles.']]
        
        #############################################################################################################################################################################
        #find the last role
        #############################################################################################################################################################################
        ################### Case 1: Previous Element was a Connector ###############
        # if the previous type was a connector, find the shared role, then look at the next role
        if previous_type == 'connector':
            # Step 1
            # find the previous roles of the last connector, take the intersection of this and the current roles, this will give you the last common role
            previous_roles = list(df_connector['ElmID of Role of Connector Ends (Port)_split'][previous_idx[0]])
            last_role = list(set(previous_roles).intersection(connector_roles))
            if not last_role:
                del previous[len(previous) - 1]
                return [[element, 'Error, connector to connector do not share common role.']]
            
        ################### Case 2: Previous Element was a Port ####################
        # if the previous type was a port, validate that it was a role of the current connector
        elif previous_type == 'port':
            if previous_ID in connector_roles:
                last_role = [previous_ID]
            else:
                del previous[len(previous) - 1]
                return [[element, 'Error, previous port not a role on last connector.']]
    
        ################### Case 3: Previous Element was a Part ####################
        # if the previous type was a part, check if it was a role of the current connector, otherwise find the role who's owner is the type of the part    
        elif previous_type == 'part':
            if previous_ID in connector_roles:
                last_role = [previous_ID]
            else:
                # look at each connector in connector_roles
                for c in connector_roles:
                    # if the connector is a port, find its index and then assess if the port owner is the same as the type of the last part (meaning it is on the edge of that part)
                    if c in df_port['Elm ID'].values:
                        role_idx = df_port[df_port['Elm ID'] == c].index.tolist()
                        if df_part['Type'][previous_idx[0]] == df_port['Owner'][role_idx[0]] or df_part['Generalization_split'][previous_idx[0]] == df_port['Owner'][role_idx[0]] or  df_port['Owner'][role_idx[0]] in df_part['Generalization_split'][previous_idx[0]] :
                            last_role.append(c)
    
        ########## Case 4: Previous Element not a Connector, Part, or Port #########                      
        else:
            del previous[len(previous) - 1]
            return [[element, 'Error, no previous element from connector that is Connector, Part, or Port.']]
            
        #############################################################################################################################################################################
        #take the last role, and find the current role, then search from the current role
        #############################################################################################################################################################################
        if len(last_role) != 1:
            del previous[len(previous) - 1]
            return [[element, 'Error: Either no last role or more than one last role.']]
        #filter the list of connector roles of the last role
        current_role = list(filter(lambda a: a != last_role[0], connector_roles))
        
        #check all possible situations for "current_role"
        ########## Case 1: 0 or more than one current roles ########################
        # if the current role is empty or has more than one current role, this is an error.
        if len(current_role) != 1:
            del previous[len(previous) - 1]
            return [[element, 'Error, 0 or more than one current roles at end of connector.']]
        ########## Case 2: Current Role is a Part ##################################
        #if the current role is a part, this ends the search. it either ends as a CI or not a CI
        elif current_role[0] in df_part['Elm ID'].values:
            current_role_idx = df_part[df_part['Elm ID'] == current_role[0]].index.tolist()
            if df_part['CI Indicator (System Context)'][current_role_idx[0]] in CI_Set:
                del previous[len(previous) - 1]
                return [[element, current_role[0], 'End of path, last element is a CI.']]
            elif df_part['Is Software'][current_role_idx[0]] or df_part['Is Wire Harness'][current_role_idx[0]]:
                del previous[len(previous) - 1]
                return[[element, current_role[0], 'End of path, last element is Wire Harness or Software.']]
            else:
                del previous[len(previous) - 1]
                return [[element, current_role[0], 'End of path, last element is part, but not CI.']]
    ###################################################################################################################################################################### Solve this problem
    ###################################################################################################################################################################### Solve this problem
        ########## Case 3: Current Role is a Port ##################################
        #if the current role is a port, we look at all connectors attached to this port.
        elif current_role[0] in df_port['Elm ID'].values:
            current_role_idx = df_port[df_port['Elm ID'] == current_role[0]].index.tolist()
            potential_parts_unique = True
            #this step has two parts.  First find all of the relevant connectors from this port.  Second search all of those connectors.
            #Part 1: Find all relevant connectors from port
            #This is where things get tricky.  Connectors are uniquely identified up to their owning element
            # find out if I'm "outside looking in", meaning a connector, attached to a specific part property
            # Case: Outside Looking In
            if df_connector['Owner'][e_idx[0]] != df_port['Owner'][current_role_idx[0]]: #if this is true, then the connector owner is different from the port owner, meaning the connector is looking at a specific part property.  Note, this is not true in an unforeseen / unlikely / possibly unallowable, recursive instance in which a block owns a part typed by itself.
                ####################################################################
                #find current_part.  if on the outside, I can look to the part with port for the connector
                current_part = list(df_connector['Part with Port for Connector Ends ElmID_split'][e_idx[0]])
                print('Element: ' + str(element) + ' Current parts: ' + str(current_part))
                
                i = len(current_part) - 1
                while i >= 0:
                    print(i)
                    if current_part[i] not in df_part['Elm ID'].values: # this means that the current part with port is pointing to a non-part property.  Likely it is a reference property
                        replacement = ref_to_part(current_part[i], df_part)
                        print('Replacement: ' + str(replacement))
                        print('Length of replacement: ' + str(len(replacement)))
                        if len(replacement) == 1: current_part[i] = replacement[0]
                        else: 
                            del current_part[i]
                            i = i-1
                            continue
                    # changed this line to allow for reference properties... need to fix this bit of logic b/c its not foolproof
                    if (current_part[i] in previous) or current_part[i] in part_between(element, previous_ID, df_connector, df_part, df_port): #or (df_port['Owner'][current_role_idx[0]] != df_part['Type'][element_idx(current_part[i], df_port, df_part, df_connector)[0]] ): #or (df_part['Type'][element_idx(current_part[i], df_port, df_part, df_connector)[0]] not in df_port['Set of Types Using Port_split'][current_role_idx[0]]) : #or (df_port['Owner'][current_role_idx[0]] != df_part['Type'][element_idx(current_part[i], df_port, df_part, df_connector)[0]] ) 
                        print('Deleting current i')
                        del current_part[i]
                    i = i - 1
                print('Current parts after analysis: ' + str(current_part))
                if len(current_part) == 1:
                    if current_part[0] in df_part['Elm ID'].values:
                        current_part_idx = df_part[df_part['Elm ID'] == current_part[0]].index.tolist()
                        if df_part['CI Indicator (System Context)'][current_part_idx[0]] in CI_Set:
                            del previous[len(previous) - 1]
                            return [[element, current_part[0], 'End of path, last element is CI.']]
#'_18_5_2_4a9015d_1508265488170_465010_115349' is generating no part with port for some reason, though it is in the connectors...
                if len(current_part) != 1:
                    del previous[len(previous) - 1]
                    return [[element, 'End of path no current part - may be an error like pointing to reference property. Look at: ' + str(list(df_connector['Part with Port for Connector Ends ElmID_split'][e_idx[0]]))]]
                    
                s = df_connector['ElmID of Role of Connector Ends (Port)_split']
                shared_role_connectors_idx = list(s[s.apply(lambda i: current_role[0] in i)].index.values)

                inside_connectors = []
                outside_connectors = []
                for i in shared_role_connectors_idx:
                    if df_port['Owner'][current_role_idx[0]] == df_connector['Owner'][i]: # if owner of current role (port) = owner of ith connector, then it's "inside"
                        inside_connectors.append(df_connector['Element ID'][i]) # note that we don't delete previous ones here, b/c we're looking into a part that we haven't explored before
                    else:
                        if (df_connector['Element ID'][i] not in previous) and (current_part[0] in df_connector['Part with Port for Connector Ends ElmID_split'][i]): # if "outside connector not previously searched & the current part is one of the parts at the end of this connector
                            outside_connectors.append(df_connector['Element ID'][i])                
                
                connectors_from_port = inside_connectors #+ outside_connectors
                if not connectors_from_port:
                    if df_part['Is Software'][current_part_idx[0]] or df_part['Is Wire Harness'][current_part_idx[0]]:
                        del previous[len(previous) - 1]
                        return[[element, current_part[0], 'End of path, last part property is a Wire Harness or Software.']]
                    else:
                        del previous[len(previous) - 1]
                        return [[element, current_part[0], 'End of path, Last part property has no more connectors, not a CI.']]

            else: # if we are inside looking outside, then we must look at two sets of connectors, the inside ones and the outside ones.  The inside ones can be uniquely identified and searched; the outside ones cannot, we must ID this non-uniqueness and annotate what part property they come from subsequently *********
                ####################################################################
                # find all connectors that are "inside" that is have the same owner as the current_port and share that same port, 
                s = df_connector['ElmID of Role of Connector Ends (Port)_split']
                shared_role_connectors_idx = list(s[s.apply(lambda i: current_role[0] in i)].index.values)

                inside_connectors = []
                outside_connectors = []
                for i in shared_role_connectors_idx:
                    if df_port['Owner'][current_role_idx[0]] == df_connector['Owner'][i]: # if owner of current role (port) = owner of ith connector, then it's "inside"
                        if df_connector['Element ID'][i] not in previous:
                            inside_connectors.append(df_connector['Element ID'][i])
                    else:
                        if df_connector['Element ID'][i] not in previous:
                            outside_connectors.append(df_connector['Element ID'][i])
                
                # inefficient code
                #inside_connectors_idx = list(s[s.apply(lambda i: df_port['Owner'][current_role_idx[0]] == i)].index.values)                    
                #for i in inside_connectors_idx:
                #    if (current_role[0] in df_connector['ElmID of Role of Connector Ends (Port)_split'][i]): 
                #        inside_connectors.append(df_connector['Element ID'][i])
                ####################################################################
                # find all connectors that are "outside", that is, share the current port, but have a different owner than the port
                #outside_connectors_idx = list(s[s.apply(lambda i: df_port['Owner'][current_role_idx[0]] != i)].index.values)
                #outside_connectors = []
                #for i in outside_connectors_idx:
                    # Multi part test: 1) can't be a previously searched connector and must have a 'part with port' the same as the current_part
                #    if (not (df_connector['Element ID'][i] in previous)):
                #        outside_connectors.append(df_connector['Element ID'][i])
                ####################################################################
                # will the set of outside connectors be unique? i.e. attached to only one part property
                # we can know this if the owner of the port (a block) is the the type for only one part property in the model
                # find all part properties in model whose owner
                s = df_part['Type']
                potential_parts_idx = list(s[s.apply(lambda i: df_port['Owner'][current_role_idx[0]]== i)].index.values)
                if len(potential_parts_idx) > 1:
                    potential_parts_unique = False
                
                # change to only look at: 1) outside looking in: inside connectors, 2) inside looking out: outside_connectors
                connectors_from_port = outside_connectors #inside_connectors + outside_connectors                
    
    
            #Part 2: Search all relevant connectors from port
            if not connectors_from_port: # this means that connectors_from_port returned 0 connectors
                ###################################################################################################################################################################### Solve this problem
                # if no more connectors, find the part or block that owns this port and return that as the end
                c_parts = list(df_connector['Part with Port for Connector Ends ElmID_split'][e_idx[0]])
                #ID the potential parts for that connector
                if not c_parts:
                    del previous[len(previous) - 1]
                    return [[element, 'Error: Connector has no Parts']]
                    
                potential_part = []
                for p in c_parts:
                    # get the index of the part
                    p_idx = df_part[df_part['Elm ID'] == p].index.tolist()
                    # if the part is typed by the same block that owns the last connector, then this is part we are looking at 
                    if df_connector['Owner'][e_idx[0]] == df_part['Type'][p_idx[0]]:
                        potential_part.append(p)
                #this is a failsafe in case the part is typed by something not in the data set
                if len(potential_part) == 0:
                    potential_part = ['Indeterminate Part Property. Last block is of type: ' + df_port['Owner'][current_role_idx[0]]]
                del previous[len(previous) - 1]
                return [[element, potential_part[0], 'End of path, connector ends at port with no more connectors']]
                
            else: #there is at least one new connector to search
                # search each connector
                for c in connectors_from_port:
                    result = search(c, previous, df_part, df_connector, df_port)
                    # if the connector was identified as being non-uniquely identifiable (meaning moving from the outside of a block that types multiple part properties)
                    if (not potential_parts_unique) and c in outside_connectors:
                        # find the index of the connector
                        c_idx = df_connector[df_connector['Element ID'] == c].index.tolist()
                        # find the parts that the connector is attached to
                        c_parts = list(df_connector['Part with Port for Connector Ends ElmID_split'][c_idx[0]])
                        #ID the potential parts for that connector
                        potential_part = []
                        for p in c_parts:
                            # get the index of the part
                            p_idx = df_part[df_part['Elm ID'] == p].index.tolist()
                            # if the part is typed by the same block that owns the last connector, then this is part we are looking at 
                            if df_connector['Owner'][e_idx[0]] == df_part['Type'][p_idx[0]]:
                                potential_part.append(p)
                        #this is a failsafe in case the part is typed by something not in the data set
                        if len(potential_part) == 0:
                            potential_part = ['Indeterminate Part Property.']
                        # ammend the result in the usual way, except ID what part we were looking at it from and ID the fact that beyond this search, we are indeterminate of where we are
                        for r in result:
                            r.insert(0,potential_part[0])
                            r.insert(0, 'Indeterminate part property from last connector.')
                            r.insert(0, element)
                            Final_Result.append(r)
                    # if the connector is uniquely identifiable, append the results in the usual way
                    else: 
                        for r in result:
                            r.insert(0, element)
                            Final_Result.append(r)
    
                # ammend previous results                            
                del previous[len(previous) - 1]
                return Final_Result  
        #if the current role is not a port or a part (may be an actor or value or something similar, for which this search is not set up)   
        ########## Case 4: Current Role is not port or a part ######################
        else:
            del previous[len(previous) - 1]
            return [[element, current_role[0], 'End of Path: Last role of connector not a port or a part.']]

################################################################################            
    # Condition 5: Element is a port
    elif e_type == 'port':
        del previous[len(previous) - 1]
        return [[element, 'End Path: Searching a port, error']]

################################################################################            
    # Condition 6: Element not a part, port, or connector  
    else:
        del previous[len(previous) - 1]
        return [[element, 'End Path: Error']]