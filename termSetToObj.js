/**
 * Term Set
 */
var TermSet = (termGroupName) => {
    /**
     * Methods
     */
    // Method to add a term to the term set object
    let addTermToSet = (termSet, termInfo, path) => {
        let term = termSet;
        let termName = "";
        // Loop for each path
        while (path.length > 0) {
            // Ensure the term exists
            termName = path[0];
            if (term[termName] == null) {
                term[termName] = {};
            }
            // Set the term
            term = term[termName];
            // Remove the term from the path
            path.splice(0, 1);
        }
        // Set the term information
        term.info = termInfo;
    };
    // Method to convert the term information into an object
    let convertToObject = (termInformation) => {
        let termSet = {};
        // Parse the terms
        for (let i = 0; i < termInformation.length; i++) {
            let term = termInformation[i];
            // Add the term to the term set
            this.addTermToSet(termSet, term, term.path.split("/"));
            // Return the term set
            return termSet;
        }
    };
    // Method to get the term information
    let getTermInfo = (terms) => {
        let termInformation = [];
        // Parse the terms
        let enumerator = terms.getEnumerator();
        while (enumerator.moveNext()) {
            let term = enumerator.get_current();
            // Add the term information
            termInformation.push({
                description: term.get_description(),
                id: term.get_id().toString(),
                name: term.get_name(),
                path: term.get_pathOfTerm().replace(/;/g, "/"),
                props: term.get_customProperties()
            });
        }
        // Sort the terms
        termInformation.sort((a, b) => {
            if (a < b) {
                return -1;
            }
            if (a > b) {
                return 1;
            }
            return 0;
        });
        // Return the terms
        return terms;
    };
    // Method to load the terms
    let loadTerms = (termGroupName) => {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the taxonomy session
            let context = SP.ClientContext.get_current();
            let session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
            // Get the term group from the default site collection term store
            let termGroup = session.getDefaultSiteCollectionTermStore().getSiteCollectionGroup(context.get_site());
            let terms = termGroup.get_termSets().getByName(termGroupName).getAllTerms();
            // Load the terms and execute the request
            context.load(terms, "Include(CustomProperties, Description, Id, Name, PathOfTerm)");
            context.executeQueryAsync(
            // Success
            () => {
                // Get the term information
                let termInfo = this.getTermInfo(terms);
                // Convert the terms to an object and resolve the promise
                resolve(convertToObject(terms));
            }, 
            // Error
            (ex, msg) => {
                // Log the error
                console.log("Error getting the terms: " + msg.get_message());
            });
        });
    };
    // Return a promise
    return new Promise((resolve, reject) => {
        SP.SOD.executeFunc("sp.js", "sp.Utilities.Utility", () => {
            // Ensure the taxonomy script is loaded
            SP.SOD.registerSOD("sp.taxonomy.js", SP.Utilities.Utility.getLayoutsPageUrl("sp.taxonomy.js"));
            SP.SOD.executeFunc("sp.taxonomy.js", "SP.Taxonomy.TaxonomySession", () => {
                // Load the terms
                this.loadTerms(termGroupName).then(termSet => {
                    // Resolve the promise
                    resolve(termSet);
                });
            });
        });
    });
};
