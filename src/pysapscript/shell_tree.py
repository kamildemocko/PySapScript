from dataclasses import dataclass

import win32com.client

from pysapscript.types_ import exceptions


@dataclass
class Node:
    _shell_tree: win32com.client.CDispatch
    key: str
    label: str
    is_expanded: bool
    is_disabled: bool
    is_folder: bool
    children_count: int

    def get_children(self) -> "list[Node]":
        if not self.is_folder:
            raise exceptions.ActionException(
                f"node with key: {self.key}, label: {self.label} has no children"
            )

        nodes = self._shell_tree.GetAllNodeKeys()
        correct_dir = False
        children = []

        for key in nodes:
            node = ShellTree._parse_node_from_list_of_nodes(self._shell_tree, key)

            if node.is_folder and key != self.key:
                correct_dir = False
                continue

            if node.is_folder and key == self.key:
                correct_dir = True
                continue
            
            if correct_dir:
                children.append(node)

        return children

    def select(self) -> None:
        self._shell_tree.SelectNode(self.key)

    def unselect(self) -> None:
        self._shell_tree.UnselectNode(self.key)

    def expand(self) -> None:
        self._shell_tree.ExpandNode(self.key)

    def collapse(self) -> None:
        self._shell_tree.CollapseNode(self.key)

    def double_click(self) -> None:
        ...


class ShellTree:
    """
    A class representing a shell table
    """

    def __init__(self, session_handle: win32com.client.CDispatch, element: str) -> None:
        """
        Usually table contains a list that can be selected and clicked

        Args:
            session_handle (win32com.client.CDispatch): SAP session handle
            element (str): SAP table element

        Raises:
            ActionException: error reading tree data
        """
        self.tree_element = element
        self._session_handle = session_handle
        self._nodes = self._read_shell_tree()

    def __repr__(self) -> str:
        return repr(self._nodes)

    def __str__(self) -> str:
        return str(self._nodes)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, ShellTree):
            return self._nodes == other._nodes
        else:
            raise NotImplementedError(f"Cannot compare ShellTable with {type(other)}")

    def __hash__(self) -> int:
        return hash(f"{self._session_handle}{self.tree_element}{len(self._nodes)}")

    def __getitem__(self, item: object) -> list[Node] | Node:
        """
        Get a node or slice of nodes from the tree.
        
        Args:
            item: An integer index or a slice object
            
        Returns:
            A single Node when indexed with an integer, or
            a list of Nodes when indexed with a slice
            
        Raises:
            ValueError: If the item type is not an integer or slice
            IndexError: If index is out of range
        """
        if isinstance(item, int):
            return self._nodes[item]

        elif isinstance(item, slice):
            return self._nodes[item.start:item.stop:item.step]

        else:
            raise ValueError("Incorrect type of index")
    
    def __len__(self) -> int:
        return len(self._nodes)

    @staticmethod
    def _parse_node_from_list_of_nodes(
        shell: win32com.client.CDispatch,
        key: str,
    ) -> Node:
        label = shell.GetNodeTextByKey(key)
        expandable = shell.IsFolderExpandable(key)
        expanded = shell.IsFolderExpanded(key)
        disabled = shell.GetIsDisabled(key, "Text")
        folder = expanded or expandable
        children_count = shell.GetNodeChildrenCount(key) if folder else 0

        return Node(
            _shell_tree=shell, 
            key=key, 
            label=label, 
            is_expanded=expanded, 
            is_disabled=disabled,
            is_folder=folder,
            children_count=children_count,
        )

    def _read_shell_tree(self) -> list[Node]:
        content = []

        shell = self._session_handle.findById(self.tree_element)
        nodes = shell.GetAllNodeKeys()

        for key in nodes:
            content.append(self._parse_node_from_list_of_nodes(shell, key))
        
        return content

    def get_node_by_key(self, key: str) -> Node | None:
        key_match = [n for n in self._nodes if n.key == key]
        if not key_match:
            return None
        
        return key_match[0]

    def get_node_by_label(self, label: str) -> Node | None:
        label_match = [n for n in self._nodes if n.label == label]
        if not label_match:
            return None
        
        return label_match[0]
    
    def get_nodes(self) -> list[Node]:
        return self._nodes
    
    def get_node_folders(self) -> list[Node]:
        return [n for n in self._nodes if n.is_folder]

    def get_node_not_folders(self) -> list[Node]:
        return [n for n in self._nodes if not n.is_folder]

    def select_all(self) -> None:
        for node in self.get_node_not_folders():
            node.select()

    def unselect_all(self) -> None:
        shell = self._session_handle.findById(self.tree_element)
        shell.UnselectAll()

    def expand_all(self) -> None:
        for folder in self.get_node_folders():
            folder.expand()

    def collapse_all(self) -> None:
        for folder in self.get_node_folders():
            folder.collapse()
