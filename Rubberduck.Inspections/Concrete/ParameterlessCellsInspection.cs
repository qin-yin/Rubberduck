using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ParameterlessCellsInspection : InspectionBase
    {
        public ParameterlessCellsInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint) { }

        public override CodeInspectionType InspectionType { get; } = CodeInspectionType.MaintainabilityAndReadabilityIssues;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var testVar = State.DeclarationFinder.MatchName("Cells");
            return null;
        }
    }
}
